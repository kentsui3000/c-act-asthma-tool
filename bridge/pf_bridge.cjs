'use strict';
// Separate worker: all clinical reads/writes go through the running Overlay API.
// No DuckDB/DBF connection and no browser or patient-identifying logs.
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

class BridgeError extends Error { constructor(code) { super(code); this.code = code; } }
const fail = code => { throw new BridgeError(code); };
function chart(value) { const s=String(value ?? ''); if(!/^\d{1,7}$/.test(s)||Number(s)===0)fail('INVALID_CHART');return s.padStart(7,'0'); }
function day(value) { const s=String(value??''); if(!/^\d{4}-\d{2}-\d{2}$/.test(s)||new Date(s+'T00:00:00Z').toISOString().slice(0,10)!==s)fail('INVALID_DATE');return s; }
function score(value) { const n=Number(value);if(value===''||!Number.isInteger(n)||n<50||n>900)fail('INVALID_PF');return n; }
function eventId(value) { if(typeof value!=='string'||! /^[A-Za-z0-9_-]{8,100}$/.test(value))fail('INVALID_EVENT_ID');return value; }
function positiveId(value) { const n=Number(value); if(!Number.isSafeInteger(n)||n<1)fail('INVALID_OVERLAY_ID');return n; }
function stamp(value) { const s=String(value??'').trim().replace(' ','T');const z=/([zZ]|[+-]\d{2}:\d{2})$/.test(s)?s:s+'+08:00';if(!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/i.test(z)||!Number.isFinite(Date.parse(z)))fail('INVALID_COMMIT_TIME');return z; }
const projection="id,kcstmr,CAST(assessment_date AS VARCHAR) AS assessment_date,score,source,entered_by,CAST(created_at AS VARCHAR) AS created_at";
const sentinel="SELECT 'probe' AS kind,NULL::BIGINT AS id,NULL::VARCHAR AS kcstmr,NULL::VARCHAR AS assessment_date,NULL::DOUBLE AS score,NULL::VARCHAR AS source,NULL::VARCHAR AS entered_by,NULL::VARCHAR AS created_at UNION ALL SELECT 'event',id,kcstmr,assessment_date,score,source,entered_by,created_at FROM (";
function exportSQL(cursor,since,limit=25) { if(!Number.isSafeInteger(cursor)||cursor<0||!Number.isSafeInteger(limit)||limit<1||limit>100)fail('INVALID_CURSOR');return sentinel+`SELECT ${projection} FROM clinical_assessments WHERE assessment_type='PF' AND source='manual' AND entered_by='tray_panel' AND superseded_by IS NULL AND id>${cursor} AND created_at>='${day(since)}'::TIMESTAMP ORDER BY id LIMIT ${limit}) q`; }
function committedSQL(id) { return sentinel+`SELECT ${projection} FROM clinical_assessments WHERE assessment_type='PF' AND entered_by='act_pf_sync:${eventId(id)}' ORDER BY id LIMIT 3) q`; }
function patientSQL(kc) { return `SELECT 'probe' AS kind,NULL::VARCHAR AS kcstmr UNION ALL SELECT 'patient',KCSTMR FROM patients WHERE KCSTMR='${chart(kc)}' LIMIT 3`; }
function probed(response) { const rows=response?.rows;if(!Array.isArray(rows)||rows.filter(r=>r.kind==='probe').length!==1||rows.some(r=>!['probe','event','patient'].includes(r.kind)))fail('DB_READ_UNCONFIRMED');return rows.filter(r=>r.kind!=='probe'); }
function mobileEvent(row) { return {eventId:eventId(row.eventId),chartNumber:chart(row.chartNumber),peakFlow:score(row.peakFlow),assessmentDate:day(row.assessmentDate)}; }
function exactCommit(rows,event) { if(rows.length>1)fail('MULTIPLE_COMMITTED_ROWS');if(!rows.length)return null;const r=rows[0];if(chart(r.kcstmr)!==event.chartNumber||day(r.assessment_date)!==event.assessmentDate||score(r.score)!==event.peakFlow||r.source!=='manual'||r.entered_by!=='act_pf_sync:'+event.eventId)fail('COMMITTED_PAYLOAD_CONFLICT');return {overlayId:positiveId(r.id),committedAt:stamp(r.created_at)}; }
function acknowledgement(response,expected,status) { const r=response?.record||response?.pfRecord||response?.event||response;if(response?.ok!==true||!r||r.eventId!==expected.eventId||chart(r.chartNumber)!==expected.chartNumber||score(r.peakFlow)!==expected.peakFlow||day(r.assessmentDate)!==expected.assessmentDate)fail('CLOUD_ACK_MISMATCH');if(status&&r.syncStatus!==status)fail('CLOUD_STATUS_MISMATCH');return r; }

class Bridge {
  // overlay.query(sql), overlay.postPF(payload), cloud.call(mode,fields), store.load/save.
  constructor({overlay,cloud,store,clock=()=>new Date(),sleep=ms=>new Promise(r=>setTimeout(r,ms)),lookbackDays=31,batchSize=25,commitAttempts=5}) {
    this.overlay=overlay;this.cloud=cloud;this.store=store;this.clock=clock;this.sleep=sleep;this.lookbackDays=lookbackDays;this.batchSize=batchSize;this.commitAttempts=commitAttempts;this.running=false;
  }
  async cycle() {
    if(this.running)fail('CYCLE_IN_PROGRESS');this.running=true;
    const result={startedAt:this.clock().toISOString(),exported:0,imported:0,rejected:0,pending:0,errorCode:null};
    try {
      this.state=await this.store.load()||{schemaVersion:1,outCursor:0};
      if(this.state.schemaVersion!==1||!Number.isSafeInteger(this.state.outCursor)||this.state.outCursor<0)fail('STATE_INVALID');
      const since=this.state.bootstrapSince||new Date(this.clock().getTime()-(this.lookbackDays-1)*86400000+8*3600000).toISOString().slice(0,10);
      day(since);
      if(!this.state.bootstrapSince){this.state.bootstrapSince=since;await this.store.save(this.state);}
      const outgoing=probed(await this.overlay.query(exportSQL(this.state.outCursor,since,this.batchSize)));
      // Validate the whole page before advancing any checkpoint; a malformed
      // or reordered response must not strand a lower id behind the cursor.
      let previous=this.state.outCursor;
      for(const row of outgoing){const id=positiveId(row.id);if(id<=previous)fail('OUTBOUND_ORDER_INVALID');previous=id;if(row.source!=='manual'||row.entered_by!=='tray_panel')fail('OUTBOUND_SOURCE_INVALID');chart(row.kcstmr);score(row.score);day(row.assessment_date);stamp(row.created_at);}
      for(const row of outgoing) {
        const id=positiveId(row.id);if(id<=this.state.outCursor)fail('OUTBOUND_ORDER_INVALID');
        if(row.source!=='manual'||row.entered_by!=='tray_panel')fail('OUTBOUND_SOURCE_INVALID');
        const e={eventId:'overlay:clinical_assessments:'+id,sourceEventId:String(id),chartNumber:chart(row.kcstmr),peakFlow:score(row.score),assessmentDate:day(row.assessment_date),recordedAt:stamp(row.created_at)};
        const ack=acknowledgement(await this.cloud.call('pf_sync',e),e,'synced');
        if(positiveId(ack.overlayId)!==id)fail('CLOUD_OVERLAY_ID_MISMATCH');
        this.state.outCursor=id;await this.store.save(this.state);result.exported++;
      }
      const pending=await this.cloud.call('pf_events',{afterRow:1,limit:this.batchSize});
      if(pending?.ok!==true||!Array.isArray(pending.events))fail('PENDING_PROTOCOL_ERROR');
      result.pending=pending.events.length;
      for(const raw of pending.events) {
        if(raw.source!=='mobile_pf'||raw.syncStatus!=='pending')fail('PENDING_SOURCE_INVALID');
        const e=mobileEvent(raw);
        let committed=exactCommit(probed(await this.overlay.query(committedSQL(e.eventId))),e);
        if(!committed) {
          const patients=probed(await this.overlay.query(patientSQL(e.chartNumber)));
          if(patients.length>1||patients.some(p=>p.kind!=='patient'||chart(p.kcstmr)!==e.chartNumber))fail('PATIENT_LOOKUP_INVALID');
          if(!patients.length) {
            acknowledgement(await this.cloud.call('pf_ack',{eventId:e.eventId,status:'rejected_patient',errorCode:'PATIENT_NOT_FOUND'}),e,'rejected_patient');result.rejected++;continue;
          }
          // This endpoint only acknowledges enqueueing. Its response is never proof of commit.
          try { await this.overlay.postPF({patient_id:'CHART-'+String(Number(e.chartNumber)),kcstmr:e.chartNumber,value_lpm:e.peakFlow,assessment_date:e.assessmentDate,entered_by:'act_pf_sync:'+e.eventId,notes:'ACT mobile PF back-entry'}); } catch (_) { /* An uncertain enqueue is resolved by exact readback below. */ }
          for(let attempt=0;attempt<this.commitAttempts&&!committed;attempt++) {
            if(attempt)await this.sleep(1000);
            committed=exactCommit(probed(await this.overlay.query(committedSQL(e.eventId))),e);
          }
          if(!committed)fail('OVERLAY_COMMIT_PENDING');
        }
        const ack=acknowledgement(await this.cloud.call('pf_ack',{eventId:e.eventId,status:'committed_overlay',...committed}),e,'synced');
        if(positiveId(ack.overlayId)!==committed.overlayId)fail('CLOUD_OVERLAY_ID_MISMATCH');result.imported++;
      }
      result.finishedAt=this.clock().toISOString();return result;
    } catch(error) { result.errorCode=error instanceof BridgeError?error.code:'TRANSPORT_FAILURE';result.finishedAt=this.clock().toISOString();return result; }
    finally {this.running=false;}
  }
}

class FileStore {
  constructor(filename){this.filename=filename;}
  async load(){try{return JSON.parse(await fs.promises.readFile(this.filename,'utf8'));}catch(e){if(e.code==='ENOENT')return null;throw e;}}
  async save(state){const tmp=this.filename+'.'+crypto.randomUUID()+'.tmp';await fs.promises.writeFile(tmp,JSON.stringify(state),{mode:0o600});await fs.promises.rename(tmp,this.filename);}
}
async function request(url,options={}) { const r=await fetch(url,{...options,signal:AbortSignal.timeout(20000)});if(!r.ok)fail('HTTP_'+r.status);let body;try{body=await r.json();}catch{fail('NON_JSON_RESPONSE');}return body; }
function adapters(config) {
  const overlayURL=new URL(config.overlayBase);
  if(overlayURL.protocol!=='http:'||overlayURL.hostname!=='127.0.0.1'||overlayURL.port!=='8080'||overlayURL.pathname!=='/')fail('INVALID_OVERLAY_URL');
  if(!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(config.scriptUrl)||typeof config.syncKey!=='string'||config.syncKey.length<32)fail('INVALID_CONFIG');
  const post=(endpoint,data)=>request(new URL(endpoint,overlayURL),{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  return {overlay:{query:sql=>post('/api/query',{sql}),postPF:payload=>post('/api/entries/peak-flow?force=true',payload)},cloud:{call:(mode,fields)=>request(config.scriptUrl,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'},body:new URLSearchParams({mode,syncKey:config.syncKey,...fields}).toString()})}};
}
async function main() {
  const folder=process.env.ACT_PF_BRIDGE_HOME||path.join(process.env.LOCALAPPDATA||'', 'ClinicACTPFBridge');
  const config=JSON.parse(await fs.promises.readFile(path.join(folder,'config.json'),'utf8'));
  const {overlay,cloud}=adapters(config);const net=require('node:net');
  // OS-held loopback lock disappears automatically on crash; second worker exits.
  const lock=net.createServer();await new Promise((resolve,reject)=>{lock.once('error',reject);lock.listen(18769,'127.0.0.1',resolve);});
  const bridge=new Bridge({overlay,cloud,store:new FileStore(path.join(folder,'state.json'))});
  await fs.promises.writeFile(path.join(folder,'worker.pid'),String(process.pid),{mode:0o600});
  let stopped=false;process.on('SIGINT',()=>{stopped=true;lock.close();});process.on('SIGTERM',()=>{stopped=true;lock.close();});
  while(!stopped){const health=await bridge.cycle();const tmp=path.join(folder,'health.tmp');await fs.promises.writeFile(tmp,JSON.stringify(health),{mode:0o600});await fs.promises.rename(tmp,path.join(folder,'health.json'));if(process.argv.includes('--once'))break;await new Promise(r=>setTimeout(r,30000));}lock.close();
}
module.exports={Bridge,BridgeError,FileStore,adapters,exportSQL,committedSQL,patientSQL,probed,chart,day,score,stamp,acknowledgement};
if(require.main===module)main().catch(e=>{console.error(e instanceof BridgeError?e.code:'BRIDGE_START_FAILED');process.exitCode=1;});
