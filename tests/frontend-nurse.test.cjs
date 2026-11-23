'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
class Element {
  constructor(tag='div') { this.tagName=tag; this.children=[]; this.listeners={}; this.style={}; this.value=''; this.textContent=''; this.disabled=false; this.checked=false; this.attributes={}; const classes=new Set(); this.classList={add:x=>classes.add(x),remove:x=>classes.delete(x),toggle:(x,on)=>on?classes.add(x):classes.delete(x)}; }
  appendChild(child) {this.children.push(child);return child;}
  removeChild(child) {this.children.splice(this.children.indexOf(child),1);}
  replaceChildren(...children){this.children=children;this.textContent='';}
  addEventListener(type,callback){(this.listeners[type]??=[]).push(callback);}
  setAttribute(key,value){this.attributes[key]=value;}
  querySelectorAll(){return [];}
  focus(){}
  submit(){this.submitted=true;}
  async click(){for(const listener of this.listeners.click||[]) await listener.call(this);}
}
function response(data,{status=200,type='application/json',url='https://script.googleusercontent.com/macros/echo'}={}) {
  return {ok:status>=200&&status<300,status,url,headers:{get:()=>type},text:async()=>typeof data==='string'?data:JSON.stringify(data)};
}
function harness(fetcher) {
  const nodes=new Map([...html.matchAll(/\bid="([^"]+)"/g)].map(m=>[m[1],new Element()]));
  const timers=new Map();let timerId=0;const calls=[];
  const document={body:new Element('body'),getElementById(id){assert.ok(nodes.has(id),'Missing DOM id '+id);return nodes.get(id);},createElement:tag=>new Element(tag),querySelectorAll:()=>[],querySelector:()=>null};
  const context={document,AbortController,URLSearchParams,Promise,console:{error(){},warn(){},log(){}},Date,
    alert(){},prompt(){context.prompts++;return 'test-pin';},prompts:0,
    setTimeout(callback,ms){const id=++timerId;timers.set(id,{callback,ms});return id;},clearTimeout(id){timers.delete(id);},
    fetch(url,options){if(!options || options.method!=='POST')return Promise.reject(new Error('Offline weather'));calls.push({url,options});return fetcher?fetcher(url,options):Promise.resolve(response({ok:true,apiVersion:2,date:'2026-11-24',rows:[]}));}};
  vm.createContext(context);
  vm.runInContext(script.replace('    })();', 'globalThis.app={nurseRequest,nurseState,loadNurseList,renderNurseRows,validPeakFlow,taipeiTimestamp,sendReportAndSave};\n    })();'),context);
  return {app:context.app,nodes,calls,timers,context};
}
function login(h){h.app.nurseState.pin='test-pin';h.app.nurseState.session=1;}
function prepareEditor(h,value='320'){login(h);h.app.nurseState.editing={rowIndex:2,rowKey:'old-key',oldValue:300};h.nodes.get('pf-modal-input').value=value;h.nodes.get('pf-modal').style.display='flex';}
const row={rowIndex:2,rowKey:'row-key',name:'Synthetic',chart:'1',testType:'C-ACT',score:0,level:'test',levelKey:'poor',peakFlow:300};

test('both served HTML copies match and the entire embedded script initializes',()=>{
  assert.equal(html,fs.readFileSync(path.join(__dirname,'..','c-act-asthma-tool','index.html'),'utf8'));
  const h=harness();
  for(const id of ['btn-nurse-mode','btn-nurse-back','btn-nurse-refresh','pf-modal-save','pf-modal-cancel'])assert.equal(h.nodes.get(id).listeners.click.length,1,id+' must have exactly one listener');
  assert.ok(!html.includes('id="peak-flow-input"'));
  assert.ok(!html.includes('var NURSE_PIN'));
});
test('one staff-entry click prompts once and sends one credentialed POST',async()=>{
  const h=harness();await h.nodes.get('btn-nurse-mode').click();await new Promise(setImmediate);
  assert.equal(h.context.prompts,1);assert.equal(h.calls.length,1);
  const request=h.calls[0];const params=new URLSearchParams(request.options.body);
  assert.equal(params.get('mode'),'nurse_list');assert.equal(params.get('pin'),'test-pin');
  assert.equal(request.options.credentials,'omit');assert.equal(request.options.redirect,'follow');assert.ok(!request.url.includes('pin='));
});
test('transport distinguishes Google auth, HTTP errors, malformed responses and backend rejection',async t=>{
  const cases=[
    ['403',response('Forbidden',{status:403,type:'text/html'}),'GOOGLE_AUTH'],
    ['auth HTML',response('<!DOCTYPE html><title>Sign in</title>',{type:'text/html'}),'GOOGLE_AUTH'],
    ['Google login redirect',response('{}',{url:'https://accounts.google.com/ServiceLogin'}),'GOOGLE_AUTH'],
    ['HTTP500',response('{}',{status:500}),'HTTP_ERROR'],
    ['bad JSON',response('{bad'),'INVALID_RESPONSE'],
    ['wrong content type',response('{}',{type:'text/plain'}),'INVALID_RESPONSE'],
    ['old backend',response({rows:[]}),'INVALID_RESPONSE'],
    ['backend rejection',response({ok:false,apiVersion:2,error:{code:'STALE_ROW',message:'Record changed',retryable:false}}),'STALE_ROW']
  ];
  for(const [label,res,code]of cases)await t.test(label,async()=>{const h=harness(()=>Promise.resolve(res));login(h);await assert.rejects(h.app.nurseRequest('nurse_list'),error=>error.code===code);assert.equal(h.timers.size,0);});
});
test('network errors and timeout fail explicitly',async()=>{
  const h=harness(()=>Promise.reject(new TypeError('Failed to fetch')));login(h);await assert.rejects(h.app.nurseRequest('nurse_list'),e=>e.code==='NETWORK_ERROR');
  const timed=harness((url,opts)=>new Promise((resolve,reject)=>opts.signal.addEventListener('abort',()=>{const err=new Error('aborted');err.name='AbortError';reject(err);})));login(timed);
  const pending=timed.app.nurseRequest('nurse_list');[...timed.timers.values()].find(t=>t.ms===20000).callback();await assert.rejects(pending,e=>e.code==='TIMEOUT');assert.equal(timed.timers.size,0);
});
test('failed list never displays an empty-day success and invalid timestamps are disclosed',async()=>{
  const h=harness(()=>Promise.resolve(response('Forbidden',{status:403})));login(h);assert.equal(await h.app.loadNurseList(),false);
  assert.match(h.nodes.get('nurse-list-container').children[0].textContent,/重新授權/);assert.doesNotMatch(h.nodes.get('nurse-list-container').children[0].textContent,/尚無測驗/);
  const invalid=harness(()=>Promise.resolve(response({ok:true,apiVersion:2,date:'2026-11-24',rows:[],invalidTimestampCount:1})));login(invalid);await invalid.app.loadNurseList();assert.match(invalid.nodes.get('nurse-list-container').children.at(-1).textContent,/日期無法判讀/);
});
test('staff rows render untrusted strings literally, retain score zero and preserve PF when opening',async()=>{
  const h=harness();login(h);const name='<img src=x onerror=alert(1)>';
  h.app.renderNurseRows([{...row,name}]);const rendered=h.nodes.get('nurse-list-container').children[0];
  assert.equal(rendered.children[0].textContent,name);assert.equal(rendered.children[0].children.length,0);assert.equal(rendered.children[2].textContent,'0');
  await rendered.click();assert.equal(h.nodes.get('pf-modal-input').value,'300');assert.equal(h.app.nurseState.editing.rowKey,'row-key');
});
test('PF requires a whole number from 50 through 900',()=>{
  const h=harness();for(const val of ['50','320','900'])assert.equal(h.app.validPeakFlow(val),true,val);
  for(const val of ['','49','901','3e2','50.5','-50','320abc','Infinity'])assert.equal(h.app.validPeakFlow(val),false,val);
});
test('PF success requires matching prior key, row identity, new key and numeric readback',async t=>{
  const base={ok:true,apiVersion:2,previousRowKey:'old-key',rowKey:'new-key',rowIndex:2,peakFlow:320};
  for(const [label,patch]of [['wrong prior key',{previousRowKey:'other'}],['missing new key',{rowKey:''}],['wrong identity',{rowIndex:3}],['string PF',{peakFlow:'320'}],['wrong PF',{peakFlow:321}]])await t.test(label,async()=>{
    const h=harness(()=>Promise.resolve(response({...base,...patch})));prepareEditor(h);await h.nodes.get('pf-modal-save').click();
    assert.match(h.nodes.get('pf-modal-status').textContent,/尚未收到一致/);assert.equal(h.nodes.get('pf-modal-input').value,'320');assert.equal(h.nodes.get('pf-modal-save').disabled,false);assert.equal(h.calls.length,1);
  });
  const h=harness((url,opts)=>Promise.resolve(response(new URLSearchParams(opts.body).get('mode')==='update_pf'?base:{ok:true,apiVersion:2,date:'2026-11-24',rows:[{...row,rowKey:'new-key',peakFlow:320}]})));
  prepareEditor(h);await h.nodes.get('pf-modal-save').click();assert.match(h.nodes.get('pf-modal-status').textContent,/已確認儲存/);assert.equal(h.app.nurseState.editing.rowKey,'new-key');assert.equal(h.app.nurseState.editing.oldValue,320);assert.equal(h.calls.length,2);assert.equal(h.nodes.get('pf-modal-save').disabled,false);
});
test('pending PF save disables repeat actions and preserves user input on failure',async()=>{
  let resolve;const h=harness(()=>new Promise(r=>resolve=r));prepareEditor(h);const pending=h.nodes.get('pf-modal-save').click();
  assert.equal(h.nodes.get('pf-modal-save').disabled,true);assert.equal(h.nodes.get('pf-modal-input').disabled,true);await h.nodes.get('pf-modal-save').click();assert.equal(h.calls.length,1);
  resolve(response({ok:false,apiVersion:2,error:{code:'STALE_ROW',message:'請重新整理',retryable:false}}));await pending;assert.equal(h.nodes.get('pf-modal-input').value,'320');assert.equal(h.nodes.get('pf-modal-save').disabled,false);assert.equal(h.app.nurseState.editing.rowKey,'old-key');
});
test('logout clears credentials and ignores late list/save responses',async()=>{
  let resolve;const h=harness(()=>new Promise(r=>resolve=r));login(h);const pending=h.app.loadNurseList();await h.nodes.get('btn-nurse-back').click();
  resolve(response({ok:true,apiVersion:2,date:'2026-11-24',rows:[row]}));await pending;assert.equal(h.app.nurseState.pin,'');assert.equal(h.nodes.get('nurse-list-container').children.length,0);
  let finish;const saving=harness(()=>new Promise(r=>finish=r));prepareEditor(saving);const save=saving.nodes.get('pf-modal-save').click();await saving.nodes.get('btn-nurse-back').click();finish(response({ok:true,apiVersion:2,previousRowKey:'old-key',rowKey:'new-key',rowIndex:2,peakFlow:320}));await save;
  assert.equal(saving.nodes.get('pf-modal-status').textContent,'');assert.equal(saving.app.nurseState.editing,null);assert.equal(saving.nodes.get('pf-modal').style.display,'none');assert.equal(saving.calls.length,1);
});
test('Taipei timestamp includes a correct explicit offset and legacy delivery is submitted-only',()=>{
  const h=harness();const timestamp=h.app.taipeiTimestamp();assert.match(timestamp,/\+08:00$/);assert.ok(Math.abs(Date.parse(timestamp)-Date.now())<1100);
  let result;h.app.sendReportAndSave('synthetic@example.invalid',(ok,message)=>result={ok,message});[...h.timers.values()].find(t=>t.ms===2500).callback();assert.equal(result.message,'submitted-only');
  assert.ok(html.includes('尚無法確認儲存與寄信'));assert.ok(!html.includes('報告已寄至您的 Email，請查收。結果已儲存至診所 Google 試算表。'));
});
