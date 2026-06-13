/* Extracted from the validated single-file build — logic unchanged. */
import { cv, normId, normName, tokens, normalizeDate } from './excelHelpers.js';

function parseFinalTemplateWorkbook(wb){
  const ws=wb.worksheets[0];
  // detect the date-header row dynamically: row with the most Date cells beyond col 5
  let headerRow=1, best=-1;
  for(let r=1;r<=Math.min(ws.rowCount,6);r++){ let n=0; for(let c=6;c<=ws.columnCount;c++){ if(cv(ws.getCell(r,c)) instanceof Date) n++; } if(n>best){best=n;headerRow=r;} }
  // map date columns
  const dateCols=[]; for(let c=1;c<=ws.columnCount;c++){ const v=cv(ws.getCell(headerRow,c)); if(v instanceof Date) dateCols.push({col:c, date:new Date(Date.UTC(v.getUTCFullYear(),v.getUTCMonth(),v.getUTCDate()))}); }
  // employee rows: after header, while col C (name) present
  const emps=[]; const firstDataRow=headerRow+1;
  for(let r=firstDataRow;r<=ws.rowCount;r++){
    const sno=cv(ws.getCell(r,1)), bio=cv(ws.getCell(r,2)), name=cv(ws.getCell(r,3));
    if((name==null||String(name).trim()==='') && (sno==null)) continue;
    if(name==null||String(name).trim()==='') continue;
    let legacy=null;
    dateCols.forEach(dc=>{ const t=cv(ws.getCell(r,dc.col)); if(typeof t==='string'){ if(/^left/i.test(t.trim()))legacy='LEFT'; else if(/^joined/i.test(t.trim()))legacy='JOINED'; } });
    emps.push({row:r, sno, id:normId(bio), bioRaw:bio, name:String(name).trim(), nname:normName(name), tokens:tokens(name), legacy,
      unit:cv(ws.getCell(r,4)), doj:normalizeDate(cv(ws.getCell(r,5))),
      openEL:numOrNull(cv(ws.getCell(r,6))), openCL:numOrNull(cv(ws.getCell(r,7))), openCO:numOrNull(cv(ws.getCell(r,8)))});
  }
  return {sheetName:ws.name, headerRow, dateCols, firstDataRow, emps,
    sumCols:{EL:9,CL:10,CO:11,ML:12,ECO:13,LOP:14,ELcf:15,CLcf:16,COcf:17,TWD:18,PAID:19}};
}
const numOrNull = v => (typeof v==='number')?v : (v!=null&&v!==''&&!isNaN(+v))?+v:null;

export { parseFinalTemplateWorkbook, numOrNull };
