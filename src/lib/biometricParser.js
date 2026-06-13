/* Extracted from the validated single-file build — logic unchanged. */
import { cv, normId, normName, tokens, normalizeDate, normalizeTime, dateKey, fmtDate } from './excelHelpers.js';
import { parsePayrollPeriodFromReportHeader } from './payrollCycle.js';

function parseEmployeeBlocks(ws, audit){
  const blocks=[]; let cur=null;
  for(let r=1;r<=ws.rowCount;r++){
    const a=cv(ws.getCell(r,1));
    if(typeof a==='string' && /^employee code/i.test(a.trim())){
      // ExcelJS returns the master value for every merged cell, so read ONE value per
      // merged region, then assign code/name by which label they follow.
      const seen=new Set(), seq=[];
      for(let c=1;c<=ws.columnCount;c++){ const cell=ws.getCell(r,c); const ad=cell.master?cell.master.address:cell.address;
        if(seen.has(ad))continue; seen.add(ad); seq.push(cv(cell)); }
      let code=null,name=null,mode=null;
      for(const v of seq){ if(v==null||v==='')continue; const s=String(v).trim();
        if(/^employee code/i.test(s)){mode='code';continue;} if(/^name/i.test(s)){mode='name';continue;}
        if(mode==='code'&&code==null)code=v; else if(mode==='name'&&name==null)name=v; }
      cur={id:normId(code), idRaw:code, name:String(name||'').trim(), nname:normName(name), tokens:tokens(name), records:[], byKey:{}};
      blocks.push(cur); continue;
    }
    if(typeof a==='string' && /^total duration/i.test(a.trim())){ cur=null; continue; }
    if(cur){
      const dateV=cv(ws.getCell(r,2)); const date=normalizeDate(dateV);
      if(!date) continue; // skip non-data rows
      const rec={ date,
        inV:cv(ws.getCell(r,4)), outV:cv(ws.getCell(r,6)),
        shift:cv(ws.getCell(r,10)), totalDur:cv(ws.getCell(r,13)),
        dayStatus:String(cv(ws.getCell(r,19))||'').trim(), remarks:String(cv(ws.getCell(r,20))||'').trim() };
      const k=dateKey(date);
      if(cur.byKey[k]!=null){ audit.push({type:'Duplicate attendance record', detail:`${cur.name} (${cur.idRaw}) has multiple rows for ${fmtDate(date)} — kept the most complete.`});
        // keep the record that has both punches
        const prev=cur.records[cur.byKey[k]];
        const score=x=>(normalizeTime(x.inV)?1:0)+(normalizeTime(x.outV)?1:0);
        if(score(rec)>score(prev)) cur.records[cur.byKey[k]]=rec;
      } else { cur.byKey[k]=cur.records.length; cur.records.push(rec); }
    }
  }
  return blocks;
}
function parseBiometricWorkbook(wb, audit){
  const ws=wb.worksheets[0];
  const period=parsePayrollPeriodFromReportHeader(ws);
  const blocks=parseEmployeeBlocks(ws, audit);
  return {sheetName:ws.name, period, blocks};
}

export { parseEmployeeBlocks, parseBiometricWorkbook };
