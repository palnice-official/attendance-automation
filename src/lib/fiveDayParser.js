/* Extracted from the validated single-file build — logic unchanged. */
import { cv, normId, normName, tokens } from './excelHelpers.js';

function parseFiveDayEmployeeList(wb){
  const ws=wb.worksheets[0]; const out=[];
  let idCol=null,nameCol=null,headerRow=null;
  for(let r=1;r<=Math.min(ws.rowCount,20);r++){
    for(let c=1;c<=ws.columnCount;c++){ const t=cv(ws.getCell(r,c));
      if(typeof t==='string'){ if(/employee\s*id/i.test(t)) {idCol=c;headerRow=r;} if(/employee\s*name/i.test(t)) nameCol=c; }
    }
    if(idCol&&nameCol){break;}
  }
  if(headerRow==null){ idCol=2;nameCol=3;headerRow=2; }
  for(let r=headerRow+1;r<=ws.rowCount;r++){
    const id=cv(ws.getCell(r,idCol)), nm=cv(ws.getCell(r,nameCol));
    if(id==null && nm==null) continue;
    out.push({id:normId(id), name:String(nm||'').trim(), nname:normName(nm), tokens:tokens(nm)});
  }
  return out;
}
function getWorkScheduleType(emp, fiveList){
  const byId = emp.id && fiveList.some(f=>f.id && f.id===emp.id);
  const byName = emp.nname && fiveList.some(f=>f.nname===emp.nname || (f.nname.length>3 && (f.nname.includes(emp.nname)||emp.nname.includes(f.nname))));
  return (byId||byName)?'5-Day':'6-Day';
}

export { parseFiveDayEmployeeList, getWorkScheduleType };
