'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),{webcrypto}=require('node:crypto');
const html=fs.readFileSync(path.join(__dirname,'..','index.html'),'utf8');
const source=html.match(/<script>([\s\S]*?)<\/script>/)[1];
class Element{
 constructor(tag='div'){this.tagName=tag;this.children=[];this.value='';this.textContent='';this.style={};this.attributes={};this.listeners={};this.checked=false;this.disabled=false;const set=new Set();this.classList={add:x=>set.add(x),remove:x=>set.delete(x),contains:x=>set.has(x),toggle:(x,on)=>on?set.add(x):set.delete(x)};}
 appendChild(child){this.children.push(child);return child;}removeChild(child){this.children.splice(this.children.indexOf(child),1);}replaceChildren(...items){this.children=items;this.textContent='';}setAttribute(k,v){this.attributes[k]=v;}querySelectorAll(){return [];}focus(){}submit(){this.submitted=true;}
 addEventListener(type,fn){(this.listeners[type]??=[]).push(fn);}async click(){if(this.disabled)return;for(const fn of this.listeners.click||[])await fn.call(this);}async input(){for(const fn of this.listeners.input||[])await fn.call(this);}
 text(){return this.textContent+this.children.map(x=>x.text()).join('\n');}
}
function response(data,status=200,type='application/json'){return{ok:status>=200&&status<300,status,url:'https://script.googleusercontent.com/macros/echo',headers:{get:()=>type},text:async()=>typeof data==='string'?data:JSON.stringify(data)}}
function history(params,extra={}){return{ok:true,apiVersion:2,startDate:params.get('startDate'),endDate:params.get('endDate'),rows:[],pfRecords:[],nextCursor:'',...extra};}
function pfFrom(params,extra={}){return{eventId:params.get('eventId'),chartNumber:params.get('chartNumber'),peakFlow:Number(params.get('peakFlow')),assessmentDate:params.get('assessmentDate'),measuredAt:'',recordedAt:'2026-11-24T02:20:00.000Z',receivedAt:'2026-11-24T02:20:00.000Z',source:'mobile_pf',sourceEventId:'',overlayId:'',syncStatus:'pending',versionKey:'version1',...extra};}
function harness(fetcher){
 const selected={};const nodes=new Map([...html.matchAll(/\bid="([^"]+)"/g)].map(m=>[m[1],new Element()])),calls=[],timers=new Map();let timerId=0,now=Date.parse('2026-11-23T17:00:00Z');
 class FakeDate extends Date{constructor(...args){super(...(args.length?args:[now]));}static now(){return now;}}
 const context={Date:FakeDate,crypto:webcrypto,Uint8Array,AbortController,URLSearchParams,Promise,console:{error(){},log(){}},alert(){},
 document:{body:new Element('body'),getElementById(id){assert.ok(nodes.has(id),'Unknown DOM ID '+id);return nodes.get(id);},createElement:tag=>new Element(tag),querySelectorAll:()=>[],querySelector:selector=>{const match=selector.match(/name="([qa]\d+)"/);return match&&selected[match[1]]!=null?{value:String(selected[match[1]])}:null;}},
 setTimeout(fn,ms){const id=++timerId;timers.set(id,{fn,ms});return id;},clearTimeout:id=>timers.delete(id),
 fetch(url,options){if(!options||options.method!=='POST')return Promise.reject(new Error('Offline weather'));const params=new URLSearchParams(options.body);calls.push({url,options,params});return fetcher?fetcher(params,options):Promise.resolve(response(history(params)));}};
 vm.createContext(context);vm.runInContext(source.replace('    })();','globalThis.app={nurseRequest,getCactTotal,getActTotal,allCactAnswered,allActAnswered,getLevelChild,getLevelAdult,sendReportAndSave,nurseState,nurseToday,normalizedChart,historyQuery,validHistoryDate,validHistoryPayload,renderNurseHistory,loadNurseList,validPeakFlow,pfPayload,savePfRecord,resetPfAttempt,pfSyncLabel,pfNeedsReview,initializeNurseDates,updatePfPatientNotice,taipeiTimestamp};\n    })();'),context);
 return{app:context.app,nodes,calls,timers,selected,advance:ms=>now+=ms};
}
function staff(h){h.app.nurseState.pin='3613';h.app.nurseState.session=1;h.app.initializeNurseDates();}
function fillPf(h,values={}){h.nodes.get('pf-chart').value=values.chart??'123';h.nodes.get('pf-date').value=values.date??'2026-11-23';h.nodes.get('pf-value').value=values.value??'95';}
const act={rowIndex:2,rowKey:'act1',chartNumber:'0000123',chart:'0000123',name:'Synthetic',testType:'C-ACT',score:0,peakFlow:'',level:'test',assessmentDate:'2026-11-24',recordedAt:'2026-11-24T01:00:00Z'};
const pf={eventId:'desk1',chartNumber:'0000123',peakFlow:95,assessmentDate:'2026-11-24',measuredAt:'',recordedAt:'2026-11-24T02:00:00Z',source:'overlay_manual',overlayId:77,syncStatus:'synced',versionKey:'v1'};


module.exports={html,source,harness,response,history,pfFrom,staff,fillPf,act,pf};
