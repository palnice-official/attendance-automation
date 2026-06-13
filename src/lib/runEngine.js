/* Extracted from the validated single-file build — logic unchanged. */
import { matchEmployees } from './matcher.js';
import { getWorkScheduleType } from './fiveDayParser.js';
import { determineAttendanceStatus, detectCompOffReviewCase, calculateLeaveUsageFromCodes, calculateLOPFromCodes, calculateClosingBalances } from './attendanceEngine.js';
import { codeMeta } from './codeDictionary.js';
import { dateKey, fmtDate, fmtTime, normalizeTime, normalizeDate, WDL } from './excelHelpers.js';
import { HOLIDAY_COVERAGE_END } from '../data/holidays.js';

function runEngine({template, biometric, fiveList, holidayMap, cycle, settings, mappingOverride, openingOverride}){
  const audit=[]; const dict=settings.dict;
  // auto match (or apply overrides)
  const auto=matchEmployees(template.emps, biometric.blocks);
  const mapping={}; // ti -> bi or null
  template.emps.forEach((t,ti)=>{ const ov=mappingOverride&&mappingOverride[ti]; mapping[ti]= ov!==undefined? ov : (auto.matches[ti]?auto.matches[ti].bi:null); });

  const dailyRows=[], missingRows=[], compRows=[], summaries=[];
  const cycleKeys=cycle.days.map(dateKey);
  let totals={present:0,lop:0,msp:0,wo:0,hol:0,comp:0,half:0,review:0};

  template.emps.forEach((t,ti)=>{
    const schedule=getWorkScheduleType(t, fiveList);
    if(schedule==='6-Day' && !(t.id && fiveList.some(f=>f.id===t.id)) && !fiveList.some(f=>f.nname===t.nname)) { /* default note added once below */ }
    const bi=mapping[ti];
    const block= (bi!=null)? biometric.blocks[bi] : null;
    const ctx={scheduleType:schedule, holidayMap, settings, dict};
    const codesForRow={}; // dateKey -> code
    const empCodes=[];

    if(!block){
      audit.push({type:'Employee on roster, missing in biometric', detail:`${t.name} (${t.bioRaw??'no id'}) — no biometric data this month; day-cells left blank.`});
      summaries.push(buildSummaryStub(t,schedule)); return;
    }
    if(auto.matches[ti] && auto.matches[ti].conf!=='hi'){
      audit.push({type:'Employee mapping (review)', detail:`${t.name} ↔ biometric "${block.name}" (${block.idRaw}) matched by ${auto.matches[ti]?auto.matches[ti].how:'override'} — please confirm.`});
    }

    cycle.days.forEach((d,di)=>{
      const k=cycleKeys[di];
      // joining handling
      if(t.doj && d < t.doj){ codesForRow[k]= (di===0||(cycle.days[di-1] && cycle.days[di-1] < t.doj))? '' : ''; // blank pre-join
        if(di===0) audit.push({type:'JOINED ON', detail:`${t.name} DOJ ${fmtDate(t.doj)} is within/after cycle start — pre-joining days left blank.`});
        return; }
      const rec=block.byKey[k]!=null? block.records[block.byKey[k]] : null;
      const st=determineAttendanceStatus(d, rec, ctx);
      codesForRow[k]=st.code; empCodes.push(st.code);

      // tallies
      const m=codeMeta(st.code,dict);
      if(m){ if(m.kind==='holiday')totals.hol++; if(m.kind==='weekoff')totals.wo++; if(m.kind==='missing')totals.msp++; if(m.kind==='half')totals.half++; if(st.code==='P')totals.present++; }
      if(st.flags.lop)totals.lop++; if(st.flags.review)totals.review++;

      // daily register row
      dailyRows.push({ id:t.bioRaw??block.idRaw, name:t.name, date:fmtDate(d), day:WDL[d.getUTCDay()], schedule,
        inT:rec?fmtTime(normalizeTime(rec.inV)):'', outT:rec?fmtTime(normalizeTime(rec.outV)):'',
        hours: st.hours!=null? st.hours.toFixed(2):'', raw: rec?rec.dayStatus:'(no record)', code:st.code, remark:st.remark, review:!!st.flags.review });

      if(st.flags.missingPunch){ missingRows.push({id:t.bioRaw??block.idRaw, name:t.name, date:fmtDate(d), day:WDL[d.getUTCDay()],
        type:`Missing ${st.flags.missingPunch.missing} Punch`, inT:rec?fmtTime(normalizeTime(rec.inV)):'', outT:rec?fmtTime(normalizeTime(rec.outV)):'', raw:rec?rec.dayStatus:'', remark:st.remark}); }
      const co=detectCompOffReviewCase(t,d,rec,st);
      if(co){ totals.comp++; compRows.push({id:t.bioRaw??block.idRaw, name:t.name, date:fmtDate(d), day:WDL[d.getUTCDay()], schedule, reason:co,
        inT:rec?fmtTime(normalizeTime(rec.inV)):'', outT:rec?fmtTime(normalizeTime(rec.outV)):'', hours:st.hours!=null?st.hours.toFixed(2):'',
        suggest:'HR to confirm CO / CO/2 / ECO / ECO/2 manually'}); }

      // legacy standalone-A audit (if any biometric ever produced it — defensive)
      if(st.code==='A'){ audit.push({type:'Standalone A detected', detail:`${t.name} ${fmtDate(d)}: standalone A → should be L for full-day unpaid absence.`}); }
    });

    const used=calculateLeaveUsageFromCodes(empCodes,dict);
    const lop=calculateLOPFromCodes(empCodes,dict);
    const ovB=(openingOverride&&openingOverride[ti])||{};
    const eff={openEL: ovB.el!=null?ovB.el:t.openEL, openCL: ovB.cl!=null?ovB.cl:t.openCL, openCO: ovB.co!=null?ovB.co:t.openCO};
    const bal=calculateClosingBalances(eff,used,settings);
    const activeDays=cycle.days.filter(d=> (!t.doj||d>=t.doj)).length;
    if(bal.clCf!=null && bal.clCf<0) audit.push({type:'Leave balance negative', detail:`${t.name}: CL closing ${bal.clCf} (<0). Excess treated as LOP per rule.`});
    if(bal.elCf<0) audit.push({type:'Leave balance negative', detail:`${t.name}: EL closing below 0 — floored to 0; excess is LOP.`});

    summaries.push({ id:t.bioRaw??block.idRaw, name:t.name, unit:t.unit||'', schedule,
      present:empCodes.filter(c=>c==='P'||c==='OD').length, half:empCodes.filter(c=>{const m=codeMeta(c,dict);return m&&m.half;}).length,
      lop, el:used.el, cl:used.cl, ml:used.ml, wo:empCodes.filter(c=>{const m=codeMeta(c,dict);return m&&m.kind==='weekoff';}).length,
      nh:empCodes.filter(c=>c==='NH').length, fh:empCodes.filter(c=>c==='FH').length, ho:empCodes.filter(c=>c==='HO').length,
      msp:empCodes.filter(c=>c==='MSP').length, co:used.co, eco:used.eco,
      openEL:eff.openEL, closeEL:bal.elCf, openCL:eff.openCL, closeCL:bal.clCf, openCO:eff.openCO, closeCO:bal.coCf,
      twd:activeDays, paid:activeDays-lop, _row:t.row, _codes:codesForRow });
  });

  // biometric-only extras -> audit
  auto.unmatchedB.forEach(bi=>{ if(Object.values(mapping).indexOf(bi)===-1){ const b=biometric.blocks[bi];
    audit.push({type:'Employee in biometric, not on roster', detail:`${b.name} (${b.idRaw}) present in biometric but not on the built-in payroll roster — not added. If a permanent hire, ask for the roster to be updated.`}); }});
  audit.push({type:'Info', detail:`Default classification: employees not on the 5-day list are treated as 6-day.`});
  audit.push({type:'Info', detail:`All Sundays marked "${settings.sundayCode}" (and Saturdays "${settings.saturdayOffCode}" for 5-day staff). Built-in holiday list applied per work schedule.`});
  // holiday-coverage warning if the cycle runs past the embedded list
  const covEnd=normalizeDate(HOLIDAY_COVERAGE_END);
  if(covEnd && cycle.end>covEnd){ audit.push({type:'Holiday coverage', detail:`This cycle extends to ${fmtDate(cycle.end)}, beyond the built-in holiday list (through ${fmtDate(covEnd)}). Jan–Mar 2027 holidays are not yet loaded — verify any holidays after ${fmtDate(covEnd)} manually.`}); }

  return {mapping, auto, dailyRows, missingRows, compRows, summaries, audit, totals};
}
function buildSummaryStub(t,schedule){ return {id:t.bioRaw??'', name:t.name, unit:t.unit||'', schedule, present:'',half:'',lop:'',el:'',cl:'',ml:'',wo:'',nh:'',fh:'',ho:'',msp:'',co:'',eco:'',openEL:t.openEL,closeEL:'',openCL:t.openCL,closeCL:'',openCO:t.openCO,closeCO:'',twd:'',paid:'',_row:t.row,_codes:null,_stub:true}; }

export { runEngine, buildSummaryStub };
