/* Extracted from the validated single-file build — logic unchanged. */
import { normalizeTime, calculateWorkingHours, parseDurationToHours } from './excelHelpers.js';
import { codeMeta } from './codeDictionary.js';
import { applyHolidayRule } from './holidayParser.js';

function determineWeekOff(date, scheduleType, settings){
  const wd=date.getUTCDay();
  if(wd===0) return {isWeekOff:true, code:settings.sundayCode};                 // Sunday
  if(wd===6 && scheduleType==='5-Day') return {isWeekOff:true, code:settings.saturdayOffCode};
  return null;
}
function validateMissingPunch(rec){
  const i=normalizeTime(rec.inV), o=normalizeTime(rec.outV);
  if(i&&!o) return {missing:'Out', cell:'out'};
  if(!i&&o) return {missing:'In', cell:'in'};
  if(/^msp$/i.test(rec.dayStatus) && !(i&&o)) return {missing: i?'Out':o?'In':'Both', cell:i?'out':'in'};
  return null;
}
function determineHalfDaySession(rec, settings){
  if(/first half/i.test(rec.dayStatus)) return {code:'P/A', remark:'First-half present (biometric)'};
  if(/second half/i.test(rec.dayStatus)) return {code:'A/P', remark:'Second-half present (biometric)'};
  return {code:settings.defaultHalfCode, remark:'Half Day - Session Requires HR Review', review:true};
}
function hasRealWork(rec, settings){
  const h=calculateWorkingHours(rec.inV,rec.outV,rec.totalDur);
  if(h!=null && h>=(settings.compOffMinHours)) return {worked:true,hours:h};
  if(/first half|second half|full day/i.test(rec.dayStatus) && (normalizeTime(rec.inV)||normalizeTime(rec.outV))) return {worked:true,hours:h};
  return {worked:false,hours:h};
}
function determineAttendanceStatus(date, rec, ctx){
  const {scheduleType, holidayMap, settings, dict} = ctx;
  const out={code:null, remark:'', flags:{}, hours:null, raw:rec?rec.dayStatus:''};
  // 1 holiday
  const hol=applyHolidayRule(date, scheduleType, holidayMap);
  if(hol){ out.code=hol.code; out.remark=hol.name;
    if(rec){ const w=hasRealWork(rec,settings); if(w.worked){ out.flags.compOff='Worked on Declared Holiday'; out.remark=`${hol.name} · Punch on Holiday - Manual Comp Off Review Required`; out.hours=w.hours; } }
    return out; }
  // 2 week-off
  const wo=determineWeekOff(date, scheduleType, settings);
  if(wo){ out.code=wo.code;
    if(rec){ const w=hasRealWork(rec,settings); if(w.worked){ const day=date.getUTCDay();
      out.flags.compOff = day===0?'Worked on Sunday Weekly Off':'Worked on Saturday Weekly Off';
      out.remark='Punch on Weekly Off - Manual Comp Off Review Required'; out.hours=w.hours; } }
    return out; }
  // (no record at all on a working day)
  if(!rec){ out.code='L'; out.remark='Full Day LOP - No Punch Available'; out.flags.lop=true; return out; }
  // 3 missing punch
  const mp=validateMissingPunch(rec);
  const i=normalizeTime(rec.inV), o=normalizeTime(rec.outV);
  if(mp){ out.code='MSP'; out.remark=`Missing ${mp.missing} Punch - Attendance Pending Verification`; out.flags.missingPunch=mp; return out; }
  // 4 both missing
  if(!i&&!o){ out.code='L'; out.remark='Full Day LOP - No Punch Available'; out.flags.lop=true; return out; }
  // 5 hours
  const h=calculateWorkingHours(rec.inV,rec.outV,rec.totalDur); out.hours=h;
  if(h==null){ out.code='MSP'; out.remark='Working hours unavailable - verify punches'; out.flags.missingPunch={missing:'Both'}; return out; }
  // pathological tiny presence flagged Absent by device -> LOP
  if(h<settings.negligibleHours && /absent/i.test(rec.dayStatus)){ out.code='L'; out.remark=`Punches present but only ${h.toFixed(2)}h and device=Absent → LOP (verify)`; out.flags.lop=true; out.flags.review=true; return out; }
  const bioHalf=/first half|second half/i.test(rec.dayStatus);
  if(settings.halfDayMode==='biometric' && bioHalf){ const s=determineHalfDaySession(rec,settings); out.code=s.code; out.remark=s.remark; if(s.review)out.flags.review=true; return out; }
  // hours-primary
  if(h>=settings.fullThreshold){ out.code='P'; out.remark='Present';
    if(bioHalf){ out.flags.review=true; out.remark=`Present (${h.toFixed(2)}h) — device flagged half-day, verify`; } return out; }
  if(h>=settings.reviewLow){ out.code='P'; out.remark='Attendance Review Required (Below 9 Hours)'; out.flags.review=true; return out; }
  // < reviewLow => half day
  const s=determineHalfDaySession(rec,settings); out.code=s.code; out.remark=`${s.remark} (${h.toFixed(2)}h)`; if(s.review)out.flags.review=true; return out;
}
function detectCompOffReviewCase(emp, date, rec, status){ return status.flags.compOff || null; }

function calculateLeaveUsageFromCodes(codes, dict){ const a={el:0,cl:0,ml:0,co:0,eco:0,present:0,half:0};
  codes.forEach(c=>{const m=codeMeta(c,dict); if(!m)return; ['el','cl','ml','co','eco','present','half'].forEach(k=>{if(m[k])a[k]+=m[k];});}); return a; }
function calculateLOPFromCodes(codes, dict){ let l=0; codes.forEach(c=>{const m=codeMeta(c,dict); if(m&&m.lop)l+=m.lop;}); return l; }
function calculateClosingBalances(emp, used, settings){
  const elCf = Math.max(0, (emp.openEL||0) + settings.elAccrual - used.el);
  const clCf = (emp.openCL!=null)? (emp.openCL - used.cl + settings.clAccrual) : null;
  const coCf = (emp.openCO||0) + used.eco - used.co;
  return {elCf, clCf, coCf};
}

export { determineWeekOff, validateMissingPunch, determineHalfDaySession, hasRealWork, determineAttendanceStatus, detectCompOffReviewCase, calculateLeaveUsageFromCodes, calculateLOPFromCodes, calculateClosingBalances };
