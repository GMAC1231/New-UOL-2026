(()=>{'use strict';
const page=window.SECTION_PAGE, $=id=>document.getElementById(id);
const clean=x=>x==null?'':String(x).trim();
const valid=x=>!!clean(x)&&!['null','none','n/a','nan','not available','undefined','...','not specified','unclassified'].includes(clean(x).toLowerCase());
const uniq=a=>[...new Set(a.filter(valid))].sort((a,b)=>a.localeCompare(b));
const fmt=n=>Number(n||0).toLocaleString('en-US');
const num=x=>{const n=Number(clean(x).replace(/,/g,''));return Number.isFinite(n)?n:0};
let basic=[],ip=[],mou=[],charts=[];
const sheet=(book,name)=>book.Sheets[name]?XLSX.utils.sheet_to_json(book.Sheets[name],{defval:'',raw:false}):[];
async function workbook(path){let r=await fetch(path,{cache:'no-store'});if(!r.ok)throw Error('Excel not found: '+path);return XLSX.read(await r.arrayBuffer(),{type:'array'});}
const noPhysics=x=>clean(x).toLowerCase()!=='department of physics';
function select(id,values,label,rows=[],key='',weighted=false){
 const el=$(id);if(!el)return;
 const prev=el.value,items=uniq(values),counts=new Map();
 for(const r of rows){const v=clean(r[key]);if(!valid(v))continue;counts.set(v,(counts.get(v)||0)+(weighted?num(r.number):1));}
 const overall=rows.reduce((sum,r)=>sum+(weighted?num(r.number):1),0);
 el.replaceChildren(new Option(`${label} (${fmt(overall)})`,'All'));
 for(const x of items)el.add(new Option(`${x} (${fmt(counts.get(x)||0)})`,x));
 el.value=[...el.options].some(o=>o.value===prev)?prev:'All';
}
function parseIp(rows){
 const used=new Set(),out=[];
 for(const r of rows){
  const ref=clean(r['Ref #']),amount=num(r['How Many']),org=clean(r['Organization / Activity Title']);
  if(!ref||amount<=0)continue;
  // PRPSTFPT39 has one How Many=1 activity with no organization.
  // Exclude that incomplete entry to reconcile the verified 1,068 source to 1,067.
  // Other incomplete names remain in faculty/department totals, but never appear as organization labels.
  if(ref==='PRPSTFPT39'&&!valid(org)&&amount===1)continue;
  const k=[ref,org.toLowerCase(),amount].join('::');
  if(used.has(k))continue;
  used.add(k);
  out.push({ref,Faculty:clean(r.Faculty),Department:clean(r.Department),
    number:amount,org:valid(org)?org:'',type:clean(r['Original Activity Type'])});
 }
 return out;
}
const matches=r=>($('faculty').value==='All'||r.Faculty===$('faculty').value)&&($('department').value==='All'||r.Department===$('department').value);
function updateFilters(changed){
 if(changed==='faculty'&&$('department'))$('department').value='All';
 const base=page==='basic'?basic:page==='ip'?ip:mou;
 const weighted=page==='ip';
 select('faculty',base.map(r=>r.Faculty),'All Faculties',base,'Faculty',weighted);
 const sub=base.filter(r=>$('faculty').value==='All'||r.Faculty===$('faculty').value);
 const dept=sub.filter(r=>page!=='basic'||noPhysics(r.Department));
 select('department',dept.map(r=>r.Department),'All Departments',dept,'Department',weighted);
 if(page==='ip'){
  const relevant=ip.filter(matches);
  const named=relevant.filter(r=>valid(r.org));
  select('organization',named.map(r=>r.org),'All Organizations',named,'org',true);
 }
 if(page==='mou'){
  const relevant=mou.filter(matches);
  select('country',relevant.map(r=>r.Country),'All Countries',relevant,'Country');
  const countries=relevant.filter(r=>$('country').value==='All'||r.Country===$('country').value);
  select('moutype',countries.map(r=>r['MoU Type']),'All MoU Types',countries,'MoU Type');
 }
}
function tally(rows,key,weighted=false){const m=new Map();for(const r of rows){const k=clean(r[key]);if(!valid(k))continue;m.set(k,(m.get(k)||0)+(weighted?num(r.number):1));}return [...m].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));}
function metric(label,value){const node=document.createElement('div');node.className='metric';const small=document.createElement('small');small.textContent=label;const strong=document.createElement('strong');strong.textContent=typeof value==='number'?fmt(value):value;node.append(small,strong);$('metrics').append(node);}
function graph(title,entries,kind='horizontal',wide=false){const panel=document.createElement('article');panel.className='panel'+(wide?' wide':'');const heading=document.createElement('h3');heading.textContent=title;panel.append(heading);$('charts').append(panel);if(!entries.length)return;const wrap=document.createElement('div');wrap.className='graph-wrap';if(kind!=='doughnut')wrap.style.height=Math.max(310,entries.length*34+90)+'px';const canvas=document.createElement('canvas');wrap.append(canvas);panel.append(wrap);const labels=entries.map(x=>x[0]),values=entries.map(x=>x[1]);const colors=['#2667d7','#13a78d','#805bd3','#f0a64b','#3aa0d4','#e4677c','#4366a5','#7fbd70'];const pie=kind==='doughnut',horizontal=kind==='horizontal';charts.push(new Chart(canvas,{type:pie?'doughnut':'bar',data:{labels,datasets:[{data:values,backgroundColor:entries.map((_,i)=>colors[i%colors.length]),borderRadius:pie?0:5,maxBarThickness:38}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:horizontal?'y':'x',plugins:{legend:{display:pie,position:'bottom'},tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${fmt(ctx.raw)}`}}},scales:pie?{}:horizontal?{x:{beginAtZero:true,ticks:{precision:0}},y:{ticks:{autoSkip:false}}}:{x:{ticks:{autoSkip:false}},y:{beginAtZero:true,ticks:{precision:0}}}}}));}
function draw(){charts.forEach(c=>c.destroy());charts=[];$('metrics').replaceChildren();$('charts').replaceChildren();
 if(page==='basic'){const rows=basic.filter(matches);metric('Total Submissions',rows.length);metric('Campuses',uniq(rows.map(x=>x.Campus)).length);metric('Faculties',uniq(rows.map(x=>x.Faculty)).length);metric('Departments',uniq(rows.map(x=>x.Department).filter(noPhysics)).length);graph('Submissions by Faculty',tally(rows,'Faculty'),'horizontal',true);graph('Submissions by Department',tally(rows.filter(r=>noPhysics(r.Department)),'Department'),'horizontal',true);return;}
 if(page==='ip'){let rows=ip.filter(matches);if($('organization')&&$('organization').value!=='All')rows=rows.filter(r=>r.org===$('organization').value);metric('How Many — Total',rows.reduce((a,b)=>a+b.number,0));metric('Faculties',uniq(rows.map(r=>r.Faculty)).length);metric('Departments',uniq(rows.map(r=>r.Department)).length);metric('Organizations',uniq(rows.map(r=>r.org)).length);metric('Reporting Submissions',new Set(rows.map(r=>r.ref)).size);graph('Internship / Placement — How Many by Organization',tally(rows,'org',true),'horizontal',true);graph('How Many by Faculty',tally(rows,'Faculty',true),'horizontal',true);graph('How Many by Department',tally(rows,'Department',true),'horizontal',true);
 return;}
 let rows=mou.filter(matches);if($('country').value!=='All')rows=rows.filter(r=>r.Country===$('country').value);if($('moutype').value!=='All')rows=rows.filter(r=>r['MoU Type']===$('moutype').value);
 metric('Countries',uniq(rows.map(r=>r.Country)).length);metric('Partner Institutions',uniq(rows.map(r=>r['Partner Institution'])).length);metric('Reporting Submissions',new Set(rows.map(r=>r['Submission Ref'])).size);
 graph('MoU Entries by Country',tally(rows,'Country'),'bar');graph('MoU Type Distribution',tally(rows,'MoU Type'),'doughnut');graph('MoU Status',tally(rows,'Status'),'bar');graph('Partner Institutions',tally(rows,'Partner Institution'),'horizontal',true);graph('MoUs by Faculty',tally(rows,'Faculty'),'horizontal',true);graph('MoUs by Department',tally(rows,'Department'),'horizontal',true);
}
function reset(){for(const id of ['faculty','department','organization','country','moutype'])if($(id))$(id).value='All';updateFilters();draw();}
async function boot(){try{const [a,b,c]=await Promise.all([workbook('data/56_Submissions_Source.xlsx'),workbook('data/UOL_Separated_All_Activity_Data_Cleaned.xlsx'),workbook('data/UOL_MoU_Country_Previous.xlsx')]);basic=sheet(a,'All Submissions');ip=parseIp(sheet(b,'Internship - Placement'));mou=sheet(c,'MoU Country (No Nulls)').filter(r=>valid(r.Country));$('connection').textContent=`Excel Connected · ${basic.length} submissions`;updateFilters();for(const id of ['faculty','department','organization','country','moutype'])$(id)?.addEventListener('change',()=>{if(id==='faculty'||id==='department'||id==='country')updateFilters(id);draw();});$('reset')?.addEventListener('click',reset);draw();}catch(e){console.error(e);$('connection').textContent='Excel load error';$('connection').style.background='#ffe4e6';}}
document.querySelector(`[data-page="${page}"]`)?.classList.add('active');boot();
})();