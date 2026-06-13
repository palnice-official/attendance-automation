/* Extracted from the validated single-file build — logic unchanged. */
function matchEmployees(templateEmps, bioBlocks){
  const tIdCount={}, bIdCount={};
  templateEmps.forEach(t=>{if(t.id)tIdCount[t.id]=(tIdCount[t.id]||0)+1;});
  bioBlocks.forEach(b=>{if(b.id)bIdCount[b.id]=(bIdCount[b.id]||0)+1;});
  const score=(t,b)=>{ let s=0; const how=[];
    if(t.id&&b.id&&t.id===b.id){ if(tIdCount[t.id]===1&&bIdCount[b.id]===1){s+=1000;how.push('ID');} else {s+=260;how.push('ID·dup');} }
    const tn=t.nname,bn=b.nname;
    if(tn&&bn){
      if(tn===bn){s+=600;how.push('name=');}
      else if(tn.length>2&&bn.length>2&&(tn.includes(bn)||bn.includes(tn))){s+=420;how.push('name~');}
      const ts=new Set(t.tokens),bs=new Set(b.tokens); const inter=[...ts].filter(x=>bs.has(x)).length; const uni=new Set([...ts,...bs]).size;
      if(uni){s+=Math.round(300*inter/uni);} if(t.tokens[0]&&t.tokens[0]===b.tokens[0])s+=120;
    }
    return {s,how:how.join('·')};
  };
  const pairs=[]; templateEmps.forEach((t,ti)=>bioBlocks.forEach((b,bi)=>{const {s,how}=score(t,b); if(s>=250)pairs.push({ti,bi,s,how});}));
  pairs.sort((a,b)=>b.s-a.s);
  const tUsed=new Set(),bUsed=new Set(),matches={};
  pairs.forEach(p=>{ if(tUsed.has(p.ti)||bUsed.has(p.bi))return; tUsed.add(p.ti);bUsed.add(p.bi);
    matches[p.ti]={bi:p.bi, score:p.s, how:p.how, conf: p.s>=900?'hi':p.s>=500?'md':'lo'}; });
  const unmatchedT=templateEmps.map((_,i)=>i).filter(i=>!tUsed.has(i));
  const unmatchedB=bioBlocks.map((_,i)=>i).filter(i=>!bUsed.has(i));
  return {matches, unmatchedT, unmatchedB};
}

export { matchEmployees };
