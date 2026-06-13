/* UI layer. Pure logic lives in ./lib; embedded assets in ./data. */
import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { DEFAULT_CODES } from './lib/codeDictionary.js';
import { pad, fmtDate, WD, b64ToArrayBuffer } from './lib/excelHelpers.js';
import { getPayrollCycle } from './lib/payrollCycle.js';
import { parseFiveDayEmployeeList } from './lib/fiveDayParser.js';
import { buildHolidayMapFromData } from './lib/holidayParser.js';
import { parseFinalTemplateWorkbook } from './lib/templateParser.js';
import { parseBiometricWorkbook } from './lib/biometricParser.js';
import { runEngine } from './lib/runEngine.js';
import { exportFinalWorkbook } from './lib/exportEngine.js';
import { TEMPLATE_B64 } from './data/template.js';
import { HOLIDAY_DATA } from './data/holidays.js';
import { RAMA_LOGO } from './data/logo.js';

const MONTHNAMES=['January','February','March','April','May','June','July','August','September','October','November','December'];
const codeClass = c => { const m=DEFAULT_CODES[c]; if(!m)return 'c-other';
  if(c==='P'||c==='OD')return 'c-P'; if(c==='L')return 'c-L'; if(m.kind==='weekoff')return 'c-WO';
  if(m.kind==='holiday')return 'c-HOL'; if(c==='MSP')return 'c-MSP'; if(m.kind==='leave')return 'c-LV';
  if(m.kind==='half')return 'c-half'; return 'c-other'; };

function Eyebrow({n,children}){ return <div className="eyebrow"><span className="num">{n}</span>{children}<span className="line"/></div>; }
function Upload({role,hint,state,onFile}){
  const ref=useRef();
  const onDrop=e=>{e.preventDefault(); if(e.dataTransfer.files[0])onFile(e.dataTransfer.files[0]);};
  return (
    <div className={'drop'+(state?.error?' err':state?.summary?' ok':'')} onClick={()=>ref.current.click()}
         onDragOver={e=>e.preventDefault()} onDrop={onDrop}>
      <div className="role">{role}</div>
      {state?.name ? <div className="fname">{state.name}</div> : <div className="hint">{hint}</div>}
      {state?.summary && <div className="meta">{state.summary}</div>}
      {state?.error && <div className="meta" style={{color:'var(--red-strong)'}}>{state.error}</div>}
      <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={e=>e.target.files[0]&&onFile(e.target.files[0])}/>
    </div>
  );
}
function Stat({v,k,tone}){ return <div className={'stat '+(tone||'')}><div className="v">{v}</div><div className="k">{k}</div></div>; }

function App(){
  const [parsed,setParsed]=useState({});            // parsed structures
  const [buffers,setBuffers]=useState({});          // raw arraybuffers
  const [year,setYear]=useState(2026), [month,setMonth]=useState(1);
  // fixed rule set (settings panel removed): Sunday=WO, hours-primary half-day, +1 EL/CL accrual
  const settings={
    fullThreshold:9.0, reviewLow:8.5, negligibleHours:1.0, compOffMinHours:4.0,
    sundayCode:'WO', saturdayOffCode:'WO', defaultHalfCode:'P/A', halfDayMode:'hours',
    elAccrual:1, clAccrual:1, cycleStartDay:25, cycleEndDay:24,
    nationalKeywords:['republic','independence','gandhi','may day','labour','labor'], dict:DEFAULT_CODES };
  const [result,setResult]=useState(null);
  const [override,setOverride]=useState({});
  const [openingOverride,setOpeningOverride]=useState({}); // ti -> {el,cl,co}
  const [tab,setTab]=useState('daily');
  const [busy,setBusy]=useState(false);
  const [err,setErr]=useState(null);
  const [tplReady,setTplReady]=useState(false);

  const cycle=useMemo(()=>getPayrollCycle(year,month,settings),[year,month]);
  const monthLabel=`${MONTHNAMES[month-1]}_${year}`;

  // load the built-in template once on mount (acts like a pre-supplied upload)
  useEffect(()=>{ (async()=>{
    try{ const buf=b64ToArrayBuffer(TEMPLATE_B64); const wb=new ExcelJS.Workbook(); await wb.xlsx.load(buf);
      const data=parseFinalTemplateWorkbook(wb);
      setParsed(p=>({...p, template:data})); setBuffers(b=>({...b, template:buf})); setTplReady(true);
    }catch(e){ setErr('Built-in template failed to load: '+(e.message||e)); }
  })(); },[]);

  const readBuf=f=>new Promise((res,rej)=>{const fr=new FileReader();fr.onload=()=>res(fr.result);fr.onerror=rej;fr.readAsArrayBuffer(f);});

  const handleFile=async(slot,f)=>{
    try{
      const buf=await readBuf(f); const wb=new ExcelJS.Workbook(); await wb.xlsx.load(buf);
      let summary='', data=null; const auditTmp=[];
      if(slot==='bio'){ data=parseBiometricWorkbook(wb,auditTmp); summary=`${data.blocks.length} employee blocks${data.period?` · ${fmtDate(data.period.from)} → ${fmtDate(data.period.till)}`:''}`;
        if(data.period&&data.period.till){ setMonth(data.period.till.getUTCMonth()+1); setYear(data.period.till.getUTCFullYear()); } }
      if(slot==='five'){ data=parseFiveDayEmployeeList(wb); summary=`${data.length} five-day employees`; }
      setBuffers(b=>({...b,[slot]:buf}));
      setParsed(p=>({...p,[slot]:data, [slot+'_state']:{name:f.name,summary}}));
      setErr(null);
    }catch(e){ setParsed(p=>({...p,[slot+'_state']:{name:f.name,error:'Could not read this file — is it a valid .xlsx?'}})); }
  };
  const ready = tplReady && parsed.bio && parsed.five;

  const process=useCallback((ovr,ovrOpen)=>{
    setBusy(true); setErr(null);
    setTimeout(()=>{
      try{
        const holidayMap=buildHolidayMapFromData(HOLIDAY_DATA, settings);
        const res=runEngine({template:parsed.template, biometric:parsed.bio, fiveList:parsed.five, holidayMap, cycle, settings,
          mappingOverride:ovr||override, openingOverride:ovrOpen||openingOverride});
        setResult(res); setBusy(false);
      }catch(e){ console.error(e); setErr(e.message||String(e)); setBusy(false); }
    },40);
  },[parsed,cycle,override,openingOverride]);

  const doExport=async()=>{ if(!result)return; setBusy(true);
    try{
      const byRow={}; if(parsed.template){ parsed.template.emps.forEach((t,ti)=>{ if(openingOverride[ti]) byRow[t.row]=openingOverride[ti]; }); }
      await exportFinalWorkbook(buffers.template, parsed.template, result, cycle, settings, monthLabel, byRow);
    }
    catch(e){ setErr('Export failed: '+(e.message||e)); } setBusy(false);
  };

  const setMap=(ti,bi)=>{ const o={...override,[ti]: bi==='__none__'?null:+bi}; setOverride(o); process(o); };
  const setOpen=(ti,field,val)=>{ const v=val===''?null:+val; const cur={...(openingOverride[ti]||{})}; cur[field]=v;
    const o={...openingOverride,[ti]:cur}; setOpeningOverride(o); };

  return (
   <div>
    <div className="topbar">
      <div className="brand">
        <img className="logo-img" src={RAMA_LOGO} alt="RAMA"/>
        <div><h1>Attendance Automation</h1><p>Biometric → payroll-ready workbook · 25th–24th cycle</p></div>
      </div>
      <button className="btn btn-export" disabled={!result||busy} onClick={doExport}>
        {busy?<span className="spinner"/>:'⬇'} Export Attendance_Final_{monthLabel}.xlsx
      </button>
    </div>

    {/* 1 UPLOAD */}
    <Eyebrow n="1">Upload source files</Eyebrow>
    <div className="grid g2">
      <Upload role="Raw biometric report" hint="Daily Attendance Detail Report" state={parsed.bio_state} onFile={f=>handleFile('bio',f)}/>
      <Upload role="5-day working employee list" hint="Mon–Fri workers (Saturdays off)" state={parsed.five_state} onFile={f=>handleFile('five',f)}/>
    </div>
    <div className="callout info" style={{marginTop:12}}>
      The payroll output template and the FY2026 holiday calendar (6-day &amp; 5-day) are <b>built in</b> — no need to upload them. Only the two files above are required each month.
    </div>

    {/* 2 MONTH + OPENING BALANCES */}
    <Eyebrow n="2">Payroll month &amp; opening balances</Eyebrow>
    <div className="card">
      <div className="grid g2" style={{gap:10,maxWidth:520}}>
        <div><label className="fld">Payroll month</label>
          <select value={month} onChange={e=>setMonth(+e.target.value)}>{MONTHNAMES.map((m,i)=><option key={i} value={i+1}>{m}</option>)}</select></div>
        <div><label className="fld">Year</label><input type="number" value={year} onChange={e=>setYear(+e.target.value)}/></div>
      </div>
      <div className="sub" style={{marginTop:12,fontWeight:600}}>Cycle: {fmtDate(cycle.start)} → {fmtDate(cycle.end)} · {cycle.days.length} days {parsed.bio_state&&!parsed.bio_state.error?'(auto-detected from biometric report)':''}</div>
      <div className="cycle">{cycle.days.map((d,i)=>{const wd=d.getUTCDay();const sat5=wd===6;
        return <div key={i} className={'day'+(wd===0?' wo':'')+(sat5?' sat5':'')}><div className="d">{pad(d.getUTCDate())}</div><div className="w">{WD[wd]}</div></div>;})}</div>
      <div className="sub" style={{marginTop:8}}><span className="chip c-WO">Sun</span> weekly off (all) · <span className="chip" style={{background:'#eef6f5',borderColor:'#cfe6e2'}}>Sat</span> off only for 5-day staff · fixed rules: ≥9h → P, 8.5–9h → P (review), &lt;8.5h → half, all unexplained absence → L, comp-off flagged only.</div>
    </div>
    <div className="card" style={{marginTop:14}}>
      <details>
        <summary>Opening leave balances ({tplReady&&parsed.template?parsed.template.emps.length:0} employees) — edit before running a later month</summary>
        <div className="callout" style={{marginTop:10}}>
          These are the built-in baseline as of <b>January 2026</b>. If you are processing a later month, update each person's opening EL / CL / CO here first — this month's closing becomes next month's opening. Blank means none on record.
        </div>
        <div className="tablewrap" style={{maxHeight:300,marginTop:10,borderRadius:8,border:'1px solid var(--line)'}}>
          <table><thead><tr><th>Employee</th><th className="mono">Open EL</th><th className="mono">Open CL</th><th className="mono">Open CO</th></tr></thead>
          <tbody>{tplReady&&parsed.template? parsed.template.emps.map((t,ti)=>(
            <tr key={ti}><td>{t.name}</td>
              <td><input type="number" step="0.5" className="map-sel" style={{width:80,fontFamily:'var(--mono)'}} defaultValue={t.openEL==null?'':t.openEL} onChange={e=>setOpen(ti,'el',e.target.value)}/></td>
              <td><input type="number" step="0.5" className="map-sel" style={{width:80,fontFamily:'var(--mono)'}} defaultValue={t.openCL==null?'':t.openCL} onChange={e=>setOpen(ti,'cl',e.target.value)}/></td>
              <td><input type="number" step="0.5" className="map-sel" style={{width:80,fontFamily:'var(--mono)'}} defaultValue={t.openCO==null?'':t.openCO} onChange={e=>setOpen(ti,'co',e.target.value)}/></td>
            </tr>)):<tr><td className="sub">Loading roster…</td></tr>}</tbody></table>
        </div>
      </details>
    </div>

    {/* 3 PROCESS */}
    <Eyebrow n="3">Process attendance</Eyebrow>
    <div style={{display:'flex',gap:12,alignItems:'center',flexWrap:'wrap'}}>
      <button className="btn btn-primary" disabled={!ready||busy} onClick={()=>process()}>{busy?<span className="spinner"/>:'▶'} Run processing</button>
      {!ready && <span className="sub">{tplReady?'Upload the biometric report and the 5-day list to enable processing.':'Loading built-in template…'}</span>}
    </div>
    {err && <div className="errbox"><b>Processing error.</b> {err}</div>}

    {result && <>
      {/* DASHBOARD */}
      <Eyebrow n="4">Processing dashboard</Eyebrow>
      <div className="grid g4">
        <Stat v={result.summaries.filter(s=>!s._stub).length} k="Employees processed" tone="good"/>
        <Stat v={result.totals.present} k="Present (P) day-cells"/>
        <Stat v={result.totals.lop} k="LOP (L) day-cells" tone={result.totals.lop?'alert':''}/>
        <Stat v={result.totals.msp} k="Missing punch (MSP)" tone={result.totals.msp?'warn':''}/>
        <Stat v={result.totals.wo} k="Weekly offs"/>
        <Stat v={result.totals.hol} k="Holiday day-cells"/>
        <Stat v={result.compRows.length} k="Comp-off review cases" tone={result.compRows.length?'warn':''}/>
        <Stat v={result.audit.length} k="Validation log entries" tone={result.audit.length>30?'warn':''}/>
      </div>

      {/* MAPPING REVIEW */}
      <Eyebrow n="5">Employee mapping review</Eyebrow>
      <div className="note">Each payroll employee is matched to a biometric block by ID, then by name. Confirm or correct any low-confidence match — changes re-process instantly. Unmatched payroll rows keep their existing template values.</div>
      <div className="tablewrap" style={{borderRadius:10,border:'1px solid var(--line)',marginTop:10}}>
        <table><thead><tr><th>Payroll employee</th><th className="mono">Tmpl ID</th><th>Confidence</th><th>Matched biometric block</th><th>Override</th></tr></thead>
        <tbody>{parsed.template.emps.map((t,ti)=>{ const bi=result.mapping[ti]; const m=result.auto.matches[ti];
          const conf= bi==null?'none': (m&&override[ti]===undefined? m.conf : 'hi');
          return <tr key={ti} className={conf==='lo'||bi==null?'row-review':''}>
            <td>{t.name}</td><td className="mono">{t.bioRaw??'—'}</td>
            <td><span className={'pill '+conf}>{ {hi:'high',md:'review',lo:'low',none:'unmatched'}[conf] }{m&&conf!=='none'&&conf!=='hi'?' · '+m.how:''}</span></td>
            <td>{bi!=null? <span>{parsed.bio.blocks[bi].name} <span className="mono sub">({parsed.bio.blocks[bi].idRaw})</span></span> : <span className="sub">— none —</span>}</td>
            <td><select className="map-sel" value={bi==null?'__none__':bi} onChange={e=>setMap(ti,e.target.value)}>
              <option value="__none__">— none —</option>
              {parsed.bio.blocks.map((b,j)=><option key={j} value={j}>{b.name} ({b.idRaw})</option>)}
            </select></td>
          </tr>;})}</tbody></table>
      </div>

      {/* PREVIEW TABS */}
      <Eyebrow n="6">Review &amp; export</Eyebrow>
      <div className="tabs">
        {[['daily','Daily Register',result.dailyRows.length],['summary','Employee Summary',result.summaries.length],
          ['missing','Missing Punch',result.missingRows.length],['comp','Comp-Off Review',result.compRows.length],
          ['audit','Audit Log',result.audit.length]].map(([k,lbl,n])=>
          <button key={k} className={'tab'+(tab===k?' active':'')} onClick={()=>setTab(k)}>{lbl}<span className="badge">{n}</span></button>)}
      </div>
      {tab==='daily' && <Tbl rows={result.dailyRows.slice(0,1500)} cols={[
        ['name','Employee'],['date','Date'],['day','Day'],['schedule','Sched'],['inT','In'],['outT','Out'],['hours','Hrs'],['raw','Raw status'],['code','Code',true],['remark','Remarks']]} reviewKey="review"/>}
      {tab==='summary' && <Tbl rows={result.summaries} cols={[
        ['name','Employee'],['unit','Unit'],['schedule','Sched'],['present','P'],['half','Half'],['lop','LOP'],['el','EL'],['cl','CL'],['ml','ML'],
        ['wo','WO'],['msp','MSP'],['closeEL','EL cf'],['closeCL','CL cf'],['closeCO','CO cf'],['twd','TWD'],['paid','Paid']]}/>}
      {tab==='missing' && <Tbl rows={result.missingRows} cols={[['name','Employee'],['date','Date'],['day','Day'],['type','Type'],['inT','In'],['outT','Out'],['raw','Raw status'],['remark','Remarks']]}/>}
      {tab==='comp' && <Tbl rows={result.compRows} cols={[['name','Employee'],['date','Date'],['day','Day'],['schedule','Sched'],['reason','Reason'],['inT','In'],['outT','Out'],['hours','Hrs'],['suggest','Suggested review']]}/>}
      {tab==='audit' && <Tbl rows={result.audit.map((a,i)=>({n:i+1,...a}))} cols={[['n','#'],['type','Category'],['detail','Detail']]}/>}

      <div className="callout info" style={{marginTop:16}}>
        The exported workbook uses the <b>built-in Sheet1 format</b> (fonts, borders, widths) with the date-column headers re-dated to the processed month, the day-codes written, and the summary formulas repaired — then appends the five report sheets. Open it in Excel; formulas recalculate on load. All leaves are written as <b>L</b> first per your rule; adjust EL/CL/ML/OD/CO manually after review.
      </div>
    </>}

    <div style={{marginTop:40,paddingTop:16,borderTop:'1px solid var(--line)'}} className="sub">
      Runs entirely in your browser · no data leaves this device · attendance code dictionary is configurable in code (DEFAULT_CODES).
    </div>
    <div className="copyright">© Copyrights owned by RAMA Group of Companies 2026.</div>
   </div>
  );
}

function Tbl({rows,cols,reviewKey}){
  if(!rows.length) return <div className="tablewrap" style={{padding:20}} ><div className="sub">No rows.</div></div>;
  return <div className="tablewrap"><table>
    <thead><tr>{cols.map(c=><th key={c[0]} className={c[2]?'mono':''}>{c[1]}</th>)}</tr></thead>
    <tbody>{rows.map((r,i)=><tr key={i} className={reviewKey&&r[reviewKey]?'row-review':''}>
      {cols.map(c=>{ const v=r[c[0]];
        if(c[0]==='code') return <td key={c[0]} className="mono">{v?<span className={'chip '+codeClass(v)}>{v}</span>:''}</td>;
        return <td key={c[0]} className={c[2]?'mono':''}>{v===null||v===undefined?'':String(v)}</td>;})}
    </tr>)}</tbody></table></div>;
}

export default App;
