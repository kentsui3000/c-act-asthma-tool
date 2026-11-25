'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const {harness,response,history,pfFrom,staff,act,pf}=require('./frontend-history-harness.cjs');

async function listed(options={}){
  const records=options.records||[act];
  const h=harness(params=>{
    if(params.get('mode')==='pf_record')return options.write?options.write(params):Promise.resolve(response({ok:true,apiVersion:2,record:pfFrom(params)}));
    return Promise.resolve(response(history(params,{rows:records,pfRecords:options.pfs||[]})));
  });
  staff(h);await h.app.loadNurseList();return h;
}
test('past single-day overview needs no chart; multi-day overview still fails closed',()=>{
  const h=harness();staff(h);h.nodes.get('nurse-start-date').value='2026-11-23';h.nodes.get('nurse-end-date').value='2026-11-23';
  assert.equal(h.app.historyQuery().chartNumber,'');assert.equal(h.app.historyQuery().clinicSession,'all');
  h.nodes.get('nurse-end-date').value='2026-11-24';assert.throws(()=>h.app.historyQuery(),e=>e.code==='CHART_REQUIRED');
  h.nodes.get('nurse-history-chart').value='123';assert.equal(h.app.historyQuery().chartNumber,'0000123');
});
test('single-day date-picker changes keep the overview on exactly that day',async()=>{
  const h=harness();staff(h);h.nodes.get('nurse-start-date').value='2026-11-22';await h.nodes.get('nurse-start-date').change();assert.equal(h.nodes.get('nurse-end-date').value,'2026-11-22');
  h.nodes.get('nurse-end-date').value='2026-11-23';await h.nodes.get('nurse-end-date').change();assert.equal(h.nodes.get('nurse-start-date').value,'2026-11-23');
});
test('ACT patient button brings chart and actual ACT date to the PF form without posting',async()=>{
  const h=await listed();const button=h.app.nurseState.quickButtons[0];assert.match(button.textContent,/Synthetic.*填入 PF/);
  await button.click();assert.equal(h.nodes.get('pf-chart').value,'0000123');assert.equal(h.nodes.get('pf-date').value,'2026-11-24');assert.equal(h.nodes.get('pf-value').value,'');
  assert.equal(h.nodes.get('pf-value').focused,true);assert.equal(h.nodes.get('pf-entry-card').scrolled,true);assert.match(h.nodes.get('pf-selected-context').textContent,/Synthetic.*0000123/);
  assert.equal(h.calls.filter(c=>c.params.get('mode')==='pf_record').length,0);
});
test('quick save keeps the clinic list, day and session and sends the verified ACT identity',async()=>{
  const h=await listed();h.nodes.get('nurse-clinic-session').value='S';await h.nodes.get('nurse-clinic-session').change();
  await h.app.nurseState.quickButtons[0].click();h.nodes.get('pf-value').value='300';assert.equal(await h.app.savePfRecord(),true);
  const write=h.calls.find(c=>c.params.get('mode')==='pf_record');assert.equal(write.params.get('linkedActIdentity'),act.stableKey);
  assert.equal(h.nodes.get('nurse-history-chart').value,'');assert.equal(h.nodes.get('nurse-clinic-session').value,'S');
  assert.equal(h.nodes.get('nurse-start-date').value,'2026-11-24');assert.equal(h.calls.at(-1).params.get('clinicSession'),'S');assert.equal(h.calls.at(-1).params.get('chartNumber'),'');
});
test('changing chart or date clears the selected ACT link instead of attaching another patient',async()=>{
  for(const [id,value] of [['pf-chart','999'],['pf-date','2026-11-23']]){
    const h=await listed();h.app.selectActForPF(act);h.nodes.get(id).value=value;await h.nodes.get(id).input();h.nodes.get('pf-value').value='300';
    assert.equal(h.app.nurseState.selectedAct,null);assert.equal(h.app.pfPayload().linkedActIdentity,undefined);
  }
});
test('unsaved PF draft prevents switching patients; an explicit clear allows the next patient',async()=>{
  const second={...act,rowKey:'act2',stableKey:'b'.repeat(64),chart:'0000999',chartNumber:'0000999',name:'Second synthetic'};
  const h=await listed({records:[act,second]});h.app.selectActForPF(act);h.nodes.get('pf-value').value='300';
  assert.equal(h.app.selectActForPF(second),false);assert.equal(h.nodes.get('pf-chart').value,'0000123');assert.equal(h.nodes.get('pf-value').value,'300');
  await h.nodes.get('btn-pf-new').click();assert.equal(h.app.selectActForPF(second),true);assert.equal(h.nodes.get('pf-chart').value,'0000999');assert.equal(h.nodes.get('pf-value').value,'');
});
test('a stale patient button cannot restore a previous result after the list changes',async()=>{
  const h=await listed();const stale=h.app.nurseState.quickButtons[0];h.app.nurseState.history.rows=[];await stale.click();
  assert.equal(h.app.nurseState.selectedAct,null);assert.equal(h.nodes.get('pf-chart').value,'');assert.match(h.nodes.get('pf-record-status').textContent,/名單已更新/);
});
test('PF save blocks selecting another patient and double submit until acknowledgment',async()=>{
  let finish;const h=await listed({write:p=>new Promise(r=>{finish=()=>r(response({ok:true,apiVersion:2,record:pfFrom(p)}));})});
  h.app.selectActForPF(act);h.nodes.get('pf-value').value='300';const save=h.app.savePfRecord();
  assert.equal(h.app.nurseState.quickButtons[0].disabled,true);assert.equal(h.app.selectActForPF(act),false);assert.equal(await h.app.savePfRecord(),false);
  finish();assert.equal(await save,true);assert.equal(h.calls.filter(c=>c.params.get('mode')==='pf_record').length,1);
});
test('wrong ACT link in acknowledgment never reports the quick entry saved',async()=>{
  const h=await listed({write:p=>Promise.resolve(response({ok:true,apiVersion:2,record:pfFrom(p,{linkedActIdentity:'b'.repeat(64)})}))});
  h.app.selectActForPF(act);h.nodes.get('pf-value').value='300';assert.equal(await h.app.savePfRecord(),false);assert.equal(h.app.nurseState.attempt.saved,false);assert.doesNotMatch(h.nodes.get('pf-record-status').textContent,/Google 已儲存/);
});
test('clinic session response and each row must match the selected filter',async()=>{
  for(const extra of [{clinicSession:'T',rows:[act]},{clinicSession:'S',rows:[{...act,clinicSession:'T'}]}]){
    const h=harness(p=>Promise.resolve(response(history(p,extra))));staff(h);h.nodes.get('nurse-clinic-session').value='S';assert.equal(await h.app.loadNurseList(),false);
  }
});
test('session change during a previous request cannot show the old clinic list',async()=>{
  const pending=[];const h=harness(p=>new Promise(r=>pending.push(()=>r(response(history(p,{rows:[{...act,clinicSession:p.get('clinicSession')}]}))))));staff(h);
  h.nodes.get('nurse-clinic-session').value='S';const first=h.app.loadNurseList();h.nodes.get('nurse-clinic-session').value='T';const second=h.app.loadNurseList();
  pending[1]();await second;pending[0]();await first;assert.match(h.nodes.get('nurse-list-date').textContent,/午診/);assert.doesNotMatch(h.nodes.get('nurse-list-date').textContent,/早診/);
});
test('date-only PF stays unassigned and only exact ACT anchors appear in the ACT supplement summary',async()=>{
  const linked={...pf,eventId:'linked1',linkedActIdentity:act.stableKey,clinicSession:'S',sessionSource:'linked_questionnaire_timestamp'};
  const other={...pf,eventId:'other1',linkedActIdentity:'b'.repeat(64)};const h=await listed({pfs:[linked,other]});
  const cards=h.nodes.get('nurse-list-container').children.filter(e=>e.tagName==='article');assert.match(cards[0].text(),/此份 ACT 補登 PF：95/);
  assert.equal((cards[0].text().match(/此份 ACT 補登 PF/g)||[]).length,1);assert.match(h.nodes.get('nurse-list-container').text(),/未分類與時段外/);
  assert.equal(pf.measuredAt,'');
});
test('chartless ACT requires chart verification and never invents a patient number',async()=>{
  const row={...act,chart:'',chartNumber:'',rowKey:'chartless'};const h=await listed({records:[row]});assert.equal(h.app.selectActForPF(row),true);
  assert.equal(h.nodes.get('pf-chart').value,'');assert.equal(h.nodes.get('pf-chart').focused,true);h.nodes.get('pf-value').value='300';assert.throws(()=>h.app.pfPayload(),e=>e.code==='INVALID_CHART');
});
test('logout clears the selected patient and delayed requests cannot reselect them',async()=>{
  const h=await listed();h.app.selectActForPF(act);await h.nodes.get('btn-nurse-back').click();assert.equal(h.app.nurseState.selectedAct,null);assert.equal(h.nodes.get('pf-chart').value,'');assert.equal(h.app.selectActForPF(act),false);
});
