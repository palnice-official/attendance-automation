/* Extracted from the validated single-file build — logic unchanged. */
import { mkUTC, cv, normalizeDate } from './excelHelpers.js';

function getPayrollCycle(year, month /*1-12, label month*/, settings){
  const sd = settings?.cycleStartDay ?? 25, ed = settings?.cycleEndDay ?? 24;
  const startMonth = month-1; // previous month index (0-based prev = month-2, but Date handles rollover)
  const start = mkUTC(year, month-2, sd);   // 25th of previous month
  const end   = mkUTC(year, month-1, ed);   // 24th of this month
  const days=[]; let d=new Date(start);
  while(d<=end){ days.push(new Date(d)); d=new Date(d.getTime()+86400000); }
  return {start,end,days};
}
function parsePayrollPeriodFromReportHeader(ws){
  for(let r=1;r<=12;r++){ for(let c=1;c<=6;c++){
    const t=cv(ws.getCell(r,c)); if(typeof t==='string' && /date from/i.test(t)){
      const m=t.match(/from:\s*([0-9A-Za-z-]+)\s*-\s*Date Till:\s*([0-9A-Za-z-]+)/i);
      if(m) return {from:normalizeDate(m[1]), till:normalizeDate(m[2])};
    }
  }}
  return null;
}

export { getPayrollCycle, parsePayrollPeriodFromReportHeader };
