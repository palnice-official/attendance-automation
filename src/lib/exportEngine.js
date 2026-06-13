/* Extracted from the validated single-file build — logic unchanged. */
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { dateKey } from './excelHelpers.js';

const colLetter = c => { let s=''; while(c>0){ s=String.fromCharCode(65+((c-1)%26))+s; c=Math.floor((c-1)/26); } return s; };
const colNum = s => { let n=0; for(const ch of String(s).toUpperCase()) n=n*26+(ch.charCodeAt(0)-64); return n; };
function fixedSummaryFormulas(r, first, last){
  const R=`${first}${r}:${last}${r}`;
  return {
    9:`=COUNTIF(${R},"EL")+COUNTIF(${R},"EL/A")/2+COUNTIF(${R},"A/EL")/2+COUNTIF(${R},"P/EL")/2+COUNTIF(${R},"EL/P")/2+COUNTIF(${R},"EL/L")/2+COUNTIF(${R},"L/EL")/2+COUNTIF(${R},"CO/EL")/2`,
    10:`=COUNTIF(${R},"CL")+COUNTIF(${R},"CL/A")/2+COUNTIF(${R},"A/CL")/2+COUNTIF(${R},"P/CL")/2+COUNTIF(${R},"CL/P")/2+COUNTIF(${R},"CL/L")/2+COUNTIF(${R},"L/CL")/2`,
    11:`=COUNTIF(${R},"CO")+COUNTIF(${R},"CO/2")/2+COUNTIF(${R},"P/CO/2")/2+COUNTIF(${R},"CO/2/P")/2+COUNTIF(${R},"CO/EL")/2`,
    12:`=IF(COUNTIF(${R},"ML")=0,"NA",COUNTIF(${R},"ML"))`,
    13:`=COUNTIF(${R},"ECO")+COUNTIF(${R},"ECO/2")/2`,
    14:`=COUNTIF(${R},"L")+COUNTIF(${R},"P/A")/2+COUNTIF(${R},"A/P")/2+COUNTIF(${R},"EL/A")/2+COUNTIF(${R},"A/EL")/2+COUNTIF(${R},"A/CL")/2+COUNTIF(${R},"CL/A")/2+COUNTIF(${R},"P/L")/2+COUNTIF(${R},"L/P")/2+COUNTIF(${R},"EL/L")/2+COUNTIF(${R},"L/EL")/2+COUNTIF(${R},"CL/L")/2+COUNTIF(${R},"L/CL")/2`,
  };
}
async function exportFinalWorkbook(templateBuffer, template, result, cycle, settings, monthLabel, openingOverridesByRow){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.load(templateBuffer);
  wb.calcProperties.fullCalcOnLoad=true;
  const ws=wb.worksheets[0];
  const dCols=template.dateCols.map(dc=>dc.col);              // 31 physical date columns (T..AX)
  const first=colLetter(dCols[0]), last=colLetter(dCols[dCols.length-1]);

  // The built-in template is dated to its original (January) cycle. Re-point the date region
  // at the month actually being processed: unmerge any date-region spans, clear old day-codes,
  // and rewrite the header dates. Columns beyond the cycle length (short months) are blanked.
  (ws.model.merges||[]).slice().forEach(mr=>{ const m=String(mr).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)/); if(!m)return;
    const c1=colNum(m[1]), c2=colNum(m[3]); if(c2>=dCols[0] && c1<=dCols[dCols.length-1]){ try{ws.unMergeCells(mr);}catch(e){} } });
  const colByKey={};
  dCols.forEach((col,i)=>{ const d=cycle.days[i];
    const hcell=ws.getCell(template.headerRow,col);
    if(d){ hcell.value=new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())); colByKey[dateKey(d)]=col; }
    else { hcell.value=null; }
    // clear every employee row's cell in this column (old month's codes)
    template.emps.forEach(e=> { ws.getCell(e.row,col).value=null; });
  });

  result.summaries.forEach(s=>{
    const r=s._row; const isStub = s._stub || !s._codes;
    // write opening balances the user edited (option a baseline) into F/G/H
    const ov=(openingOverridesByRow&&openingOverridesByRow[r])||{};
    if(ov.el!=null) ws.getCell(r,6).value=ov.el;
    if(ov.cl!=null) ws.getCell(r,7).value=ov.cl;
    if(ov.co!=null) ws.getCell(r,8).value=ov.co;
    // write daily codes for matched employees (date region is now unmerged + cleared)
    if(!isStub){
      Object.keys(s._codes).forEach(k=>{ const col=colByKey[k]; if(col){ const val=s._codes[k]; ws.getCell(r,col).value=(val===''?null:val); } });
    }
    // ALWAYS repair the summary formulas (fixes the legacy unquoted "A/EL" and short ranges,
    // so the whole workbook recalculates with zero formula errors)
    const f=fixedSummaryFormulas(r,first,last); Object.keys(f).forEach(c=>{ ws.getCell(r,+c).value={formula:f[c].slice(1)}; });
    ws.getCell(r,15).value={formula:`IF(ISNUMBER(F${r}), MAX(0, F${r} + ${settings.elAccrual} - I${r}), "NA")`}; // EL cf
    ws.getCell(r,16).value={formula:`IF(ISNUMBER(G${r}), G${r} - J${r} + ${settings.clAccrual}, "NA")`};          // CL cf
    ws.getCell(r,17).value={formula:`IF(ISNUMBER(H${r}), (H${r}+M${r})-K${r}, "NA")`};                            // CO cf
    if(!isStub) ws.getCell(r,18).value=s.twd;                                                        // TWD (matched only)
    ws.getCell(r,19).value={formula:`R${r}-N${r}`};                                                  // paid
  });

  // helper to style new report sheets consistently
  const thin={style:'thin',color:{argb:'FFDDE3EA'}};
  const addSheet=(name,cols,rows)=>{
    const sh=wb.addWorksheet(name,{views:[{state:'frozen',ySplit:1}]});
    sh.columns=cols.map(c=>({header:c.h,key:c.k,width:c.w}));
    const hr=sh.getRow(1); hr.font={name:'Times New Roman',bold:true,size:11,color:{argb:'FF0F1B2D'}};
    hr.alignment={vertical:'middle'}; hr.height=20;
    hr.eachCell(c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEDF2F1'}}; c.border={bottom:{style:'medium',color:{argb:'FF0D7D72'}}};});
    rows.forEach(rw=>{ const r=sh.addRow(rw); r.font={name:'Times New Roman',size:10.5};
      r.eachCell(c=>{c.border={bottom:thin}; c.alignment={vertical:'top'};});
      if(rw.__review){ r.eachCell(c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFDF2E1'}};}); }
      if(rw.__miss){ const idx=cols.findIndex(c=>c.k==='inT'||c.k==='outT'); }
    });
    sh.autoFilter={from:{row:1,column:1},to:{row:1,column:cols.length}};
    return sh;
  };

  // Sheet 2: Daily Attendance Register
  addSheet('Daily Attendance Register',
    [{h:'Employee ID',k:'id',w:12},{h:'Employee Name',k:'name',w:24},{h:'Date',k:'date',w:13},{h:'Day',k:'day',w:11},
     {h:'Work Schedule',k:'schedule',w:13},{h:'In Time',k:'inT',w:9},{h:'Out Time',k:'outT',w:9},{h:'Total Hours',k:'hours',w:11},
     {h:'Raw Biometric Status',k:'raw',w:20},{h:'Final Attendance',k:'code',w:13},{h:'Remarks',k:'remark',w:46}],
    result.dailyRows.map(d=>({...d, __review:d.review})));

  // Sheet 3: Employee Summary Report
  addSheet('Employee Summary Report',
    [{h:'Employee ID',k:'id',w:12},{h:'Employee Name',k:'name',w:24},{h:'Unit',k:'unit',w:18},{h:'Work Schedule',k:'schedule',w:13},
     {h:'Present',k:'present',w:9},{h:'Half Days',k:'half',w:9},{h:'LOP Days',k:'lop',w:9},{h:'EL Used',k:'el',w:8},{h:'CL Used',k:'cl',w:8},{h:'ML Used',k:'ml',w:8},
     {h:'WO',k:'wo',w:6},{h:'NH',k:'nh',w:6},{h:'FH',k:'fh',w:6},{h:'HO',k:'ho',w:6},{h:'MSP',k:'msp',w:6},{h:'CO Used',k:'co',w:8},{h:'ECO',k:'eco',w:7},
     {h:'Open EL',k:'openEL',w:8},{h:'Close EL',k:'closeEL',w:8},{h:'Open CL',k:'openCL',w:8},{h:'Close CL',k:'closeCL',w:8},{h:'Open CO',k:'openCO',w:8},{h:'Close CO',k:'closeCO',w:8},
     {h:'Total Working Days',k:'twd',w:12},{h:'Days Paid',k:'paid',w:10}],
    result.summaries);

  // Sheet 4: Missing Punch Exception Report
  addSheet('Missing Punch Exception Report',
    [{h:'Employee ID',k:'id',w:12},{h:'Employee Name',k:'name',w:24},{h:'Date',k:'date',w:13},{h:'Day',k:'day',w:11},
     {h:'Missing Punch Type',k:'type',w:18},{h:'In Time',k:'inT',w:9},{h:'Out Time',k:'outT',w:9},{h:'Raw Biometric Status',k:'raw',w:20},{h:'Remarks',k:'remark',w:46}],
    result.missingRows.map(r=>({...r,__miss:true})));

  // Sheet 5: Comp Off Review Report
  addSheet('Comp Off Review Report',
    [{h:'Employee ID',k:'id',w:12},{h:'Employee Name',k:'name',w:24},{h:'Date',k:'date',w:13},{h:'Day',k:'day',w:11},
     {h:'Work Schedule',k:'schedule',w:13},{h:'Reason',k:'reason',w:30},{h:'In Time',k:'inT',w:9},{h:'Out Time',k:'outT',w:9},{h:'Total Hours',k:'hours',w:11},{h:'Suggested Review',k:'suggest',w:42}],
    result.compRows);

  // Sheet 6: Validation / Audit Log
  addSheet('Validation _ Audit Log',
    [{h:'#',k:'n',w:5},{h:'Category',k:'type',w:34},{h:'Detail',k:'detail',w:110}],
    result.audit.map((a,i)=>({n:i+1,type:a.type,detail:a.detail})));

  const buf=await wb.xlsx.writeBuffer();
  saveAs(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), `Attendance_Final_${monthLabel}.xlsx`);
}

export { colLetter, colNum, fixedSummaryFormulas, exportFinalWorkbook };
