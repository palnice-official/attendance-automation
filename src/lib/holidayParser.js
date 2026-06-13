/* Extracted from the validated single-file build — logic unchanged. */
import { cv, normalizeDate, dateKey, fmtDate, mkUTC } from './excelHelpers.js';

function classifyHolidayType(name, settings){
  const n=String(name||'').toLowerCase();
  const nat=(settings?.nationalKeywords||['republic','independence','gandhi','may day','labour','labor']);
  if(nat.some(k=>n.includes(k))) return 'NH';
  return 'FH';
}
function parseHolidaySectionsByWorkSchedule(ws, sheetYear, audit, settings){
  // returns { '6-Day':[{date,name,type}], '5-Day':[...] }
  const res={'6-Day':[], '5-Day':[]}; let section=null;
  for(let r=1;r<=ws.rowCount;r++){
    const a=cv(ws.getCell(r,1));
    // section header detection (text may sit in col A, possibly merged)
    let line=''; for(let c=1;c<=Math.min(ws.columnCount,6);c++){const t=cv(ws.getCell(r,c)); if(typeof t==='string')line+=' '+t;}
    const L=line.toLowerCase();
    if(/5\s*day|saturdays?\s*off|with\s*saturdays/.test(L)){ section='5-Day'; continue; }
    if(/6\s*day|all\s*saturdays|working\s*on\s*all/.test(L)){ section='6-Day'; continue; }
    if(section){
      // data row: col A int sno, col B name, col C date
      const sno=cv(ws.getCell(r,1)), nm=cv(ws.getCell(r,2)), dt=cv(ws.getCell(r,3));
      if(nm && dt!=null && !/holiday/i.test(String(nm))){
        let d=normalizeDate(dt);
        if(d){
          // correct year-typos: sheet named for a year should hold that year's dates
          if(sheetYear && d.getUTCFullYear()!==sheetYear){
            audit.push({type:'Formula/Data correction', detail:`Holiday "${nm}" dated ${fmtDate(d)} in sheet "${sheetYear}" — year corrected to ${sheetYear}.`});
            d=mkUTC(sheetYear,d.getUTCMonth(),d.getUTCDate());
          }
          res[section].push({date:d, name:String(nm).trim(), type:classifyHolidayType(nm,settings)});
        }
      }
    }
  }
  return res;
}
function parseHolidayCalendar(wb, audit, settings){
  // map[schedule][dateKey] = {name,type}; reads every year-named sheet
  const map={'6-Day':{},'5-Day':{}};
  wb.worksheets.forEach(ws=>{
    const ym=ws.name.match(/(20\d{2})/); const sheetYear=ym?+ym[1]:null;
    const sec=parseHolidaySectionsByWorkSchedule(ws, sheetYear, audit, settings);
    ['6-Day','5-Day'].forEach(s=> sec[s].forEach(h=>{ map[s][dateKey(h.date)]={name:h.name,type:h.type, year:sheetYear}; }));
    if(sheetYear) audit.push({type:'Info', detail:`Holiday sheet "${ws.name}" parsed: ${sec['6-Day'].length} (6-day) + ${sec['5-Day'].length} (5-day) entries.`});
  });
  return map;
}
function applyHolidayRule(date, scheduleType, holidayMap){
  const h=holidayMap[scheduleType] && holidayMap[scheduleType][dateKey(date)];
  return h?{isHoliday:true, code:h.type, name:h.name}:null;
}

function buildHolidayMapFromData(data, settings){
  const map={'6-Day':{},'5-Day':{}};
  ['6-Day','5-Day'].forEach(s=>{ (data[s]||[]).forEach(h=>{ const d=normalizeDate(h.date); if(d) map[s][dateKey(d)]={name:h.name, type:classifyHolidayType(h.name,settings)}; }); });
  return map;
}

export { classifyHolidayType, parseHolidaySectionsByWorkSchedule, parseHolidayCalendar, applyHolidayRule, buildHolidayMapFromData };
