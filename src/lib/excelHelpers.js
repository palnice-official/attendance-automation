/* Extracted from the validated single-file build — logic unchanged. */
const MONTHS = ['january','february','march','april','may','june','july','august','september','october','november','december'];
const pad = n => String(n).padStart(2,'0');
const excelSerialToDate = s => new Date(Math.round((s-25569)*86400000));            // -> UTC date
const excelSerialToDateTime = s => new Date(Math.round((s-25569)*86400000));        // same basis, keeps time
function normalizeDate(v){
  if(v==null||v==='') return null;
  if(v instanceof Date) return new Date(Date.UTC(v.getUTCFullYear(),v.getUTCMonth(),v.getUTCDate()));
  if(typeof v==='number') { const d=excelSerialToDate(v); return new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())); }
  const s=String(v).trim();
  let m;
  if((m=s.match(/^(\d{1,2})[-/\s]([A-Za-z]{3,})[-/\s](\d{2,4})/))){ const mo=MONTHS.findIndex(x=>x.startsWith(m[2].toLowerCase())); if(mo>=0) return mkUTC(yr(m[3]),mo,+m[1]); }
  if((m=s.match(/^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})/))){ const mo=MONTHS.findIndex(x=>x===m[2].toLowerCase()); if(mo>=0) return mkUTC(+m[3],mo,+m[1]); }
  if((m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return mkUTC(+m[1],+m[2]-1,+m[3]);
  if((m=s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/))) return mkUTC(+m[3],+m[2]-1,+m[1]); // dd-mm-yyyy
  const d=new Date(s); return isNaN(d)?null:new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
}
const mkUTC=(y,mo,d)=>new Date(Date.UTC(y,mo,d));
const yr=s=>{const n=+s;return n<100?2000+n:n;};
const dateKey = d => d? `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}` : null;
const WD=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const WDL=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const fmtDate = d => d? `${pad(d.getUTCDate())}-${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getUTCMonth()]}-${d.getUTCFullYear()}` : '';
function normalizeTime(v){ // returns a Date (instant) for in/out punches, else null
  if(v==null||v==='') return null;
  if(v instanceof Date) return v;
  if(typeof v==='number') return excelSerialToDateTime(v);
  const d=new Date(String(v)); return isNaN(d)?null:d;
}
const fmtTime = d => d&&d instanceof Date&&!isNaN(d)? `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}` : '';
function parseDurationToHours(v){
  if(v==null||v==='') return null;
  if(v instanceof Date) return null;
  if(typeof v==='number') return v>0? v*24 : null;     // excel duration = fraction of a day
  const m=String(v).trim().match(/^(\d+):(\d{1,2})/); if(m){const h=+m[1]+ (+m[2])/60; return h>0?h:null;}
  const f=parseFloat(v); return (!isNaN(f)&&f>0)? f : null;
}
function calculateWorkingHours(inV,outV,durV){
  const dur=parseDurationToHours(durV); if(dur!=null) return dur;
  const i=normalizeTime(inV), o=normalizeTime(outV);
  if(i&&o&&!isNaN(i)&&!isNaN(o)){ let h=(o-i)/3600000; if(h<0) h+=24; return h; }
  return null;
}
function cv(cell){ // unwrap an ExcelJS cell value
  if(!cell) return null; let v=cell.value; if(v==null) return null;
  if(typeof v==='object'){
    if(v instanceof Date) return v;
    if('result' in v) return v.result;
    if('text' in v) return v.text;
    if('richText' in v) return v.richText.map(t=>t.text).join('');
    if('error' in v) return null;
    if('hyperlink' in v && 'text' in v) return v.text;
  }
  return v;
}
const normId = v => { if(v==null) return null; const s=String(v).trim(); if(!s) return null; return /^\d+(\.0+)?$/.test(s)?String(parseInt(s,10)):s.toUpperCase(); };
const normName = v => String(v||'').toLowerCase().replace(/[^a-z0-9]/g,'');
const tokens = v => String(v||'').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);

function b64ToArrayBuffer(b64){ const bin=atob(b64); const n=bin.length; const bytes=new Uint8Array(n); for(let i=0;i<n;i++)bytes[i]=bin.charCodeAt(i); return bytes.buffer; }

export { MONTHS, pad, excelSerialToDate, excelSerialToDateTime, normalizeDate, mkUTC, yr, dateKey, WD, WDL, fmtDate, normalizeTime, fmtTime, parseDurationToHours, calculateWorkingHours, cv, normId, normName, tokens, b64ToArrayBuffer };
