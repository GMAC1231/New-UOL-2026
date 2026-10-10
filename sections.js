(()=>{'use strict';
const page=window.SECTION_PAGE, $=id=>document.getElementById(id);
const clean=x=>x==null?'':String(x).trim();
const valid=x=>!!clean(x)&&!['null','none','n/a','nan','not available','undefined','...','not specified','unclassified'].includes(clean(x).toLowerCase());
const uniq=a=>[...new Set(a.filter(valid))].sort((a,b)=>a.localeCompare(b));
const fmt=n=>Number(n||0).toLocaleString('en-US');
const num=x=>{const n=Number(clean(x).replace(/,/g,''));return Number.isFinite(n)?n:0};
const normalizeOrg=x=>{
 let s=clean(x);
 if(!valid(s) || s==='...') return '';
 s=s.replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"');
 s=s.replace(/^[.\-_,;:]+|[.\-_,;:]+$/g,'').replace(/\s+/g,' ').trim();
 return valid(s)?s:'';
};
let basic=[],ip=[],mou=[],charts=[];
let roster=[],departmentRoster=[],facultyRoster=[],submissionByRef=new Map();
const semesterOf = value => {
 const v=clean(value).toLowerCase();
 if(!v) return 'Unassigned';
 if(v.includes('spring')) return 'Spring 2026';
 if(v.includes('fall') || v.includes('autumn')) return 'Fall 2026';
 const year=v.match(/20\d{2}/)?.[0];
 const month=v.match(/(?:20\d{2})[-/](\d{1,2})/)?.[1];
 if(year==='2026'&&month){const n=Number(month);return n<=6&&n>=1?'Spring 2026':n<=12?'Fall 2026':'Unassigned';}
 const names=['january','february','march','april','may','june','july','august','september','october','november','december'];
 const m=names.findIndex(name=>v.includes(name));
 return year==='2026' && m>=0?(m<6?'Spring 2026':'Fall 2026'):'Unassigned';
};
// Provisional semester classification from dated activities where Reporting Month is blank.
const SEMESTER_FROM_ACTIVITY_EVIDENCE = Object.freeze({
 'PRPSTFPT50':'Spring 2026',
 'PRPSTFPT48':'Spring 2026',
 'PRPSTFPT47':'Spring 2026',
 'PRPSTFPT45':'Spring 2026',
 'PRPSTFPT39':'Spring 2026',
 'PRPSTFPT38':'Spring 2026',
 'PRPSTFPT37':'Spring 2026',
 'PRPSTFPT36':'Spring 2026',
 'PRPSTFPT35':'Spring 2026',
 'PRPSTFPT26':'Spring 2026',
 'PRPSTFPT103':'Fall 2026',
 'PRPSTFPT86':'Fall 2026',
 'PRPSTFPT85':'Fall 2026',
 'PRPSTFPT82':'Fall 2026',
});

const selectedSemester=()=> $('semester')?.value||'All';
const semesterMatches=r=>selectedSemester()==='All'||r.semester===selectedSemester();
// Show the institution-wide submission counts in each section's semester menu.
// All Semesters keeps unclassified reports in totals; no unassigned filter is exposed.
function updateSemesterLabels(){
 const el=$('semester');if(!el)return;
 const counts={'All':roster.length,'Spring 2026':0,'Fall 2026':0};
 for(const row of roster){const period=semesterForSubmission(row);if(period in counts)counts[period]++;}
 for(const opt of el.options){if(opt.value in counts){
  const label=opt.value==='All'?'All Semesters':opt.value;
  opt.textContent=`${label} (${fmt(counts[opt.value])} submissions)`;
 }}
}
function semesterForSubmission(row){
 const ref=clean(row['Ref #']||row.ref||row['Submission Ref']);
 if(ref==='PRPSTFPT34')return 'Spring 2026';
 if(SEMESTER_FROM_ACTIVITY_EVIDENCE[ref])return SEMESTER_FROM_ACTIVITY_EVIDENCE[ref];
 return semesterOf(row['Reporting Month']);
}
const coverageDepartments=faculty=>uniq(departmentRoster.filter(r=>faculty==='All'||r.Faculty===faculty).map(r=>r.Department));
const coverageFaculties=()=>facultyRoster;
const rosterFacultyCount=()=>uniq(departmentRoster.filter(r=>$('faculty')?.value==='All'||r.Faculty===$('faculty')?.value).filter(r=>$('department')?.value==='All'||r.Department===$('department')?.value).map(r=>r.Faculty)).length;
const rosterDepartmentCount=()=>coverageDepartments($('faculty')?.value||'All').filter(x=>$('department')?.value==='All'||x===$('department')?.value).length;

let organizationChartLimit='15';
const sheet=(book,name)=>book.Sheets[name]?XLSX.utils.sheet_to_json(book.Sheets[name],{defval:'',raw:false}):[];
async function workbook(path){
 // Sections 1, 2 and 3 load their bundled Excel files automatically.
 const response=await fetch(path,{cache:'no-store'});
 if(!response.ok)throw Error('Unable to load '+path+' (HTTP '+response.status+')');
 return XLSX.read(await response.arrayBuffer(),{type:'array'});
}
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
  const ref=clean(r['Ref #']),amount=num(r['How Many']),org=normalizeOrg(r['Organization / Activity Title']);
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
 if((changed==='faculty'||changed==='semester')&&$('department'))$('department').value='All';
 const base=(page==='basic'?basic:page==='ip'?ip:mou).filter(semesterMatches);
 const weighted=page==='ip';
 select('faculty',facultyRoster,'All Faculties',base,'Faculty',weighted);
 const sub=base.filter(r=>$('faculty').value==='All'||r.Faculty===$('faculty').value);
 const dept=sub;
 select('department',coverageDepartments($('faculty').value),'All Departments',dept,'Department',weighted);
 if(page==='ip'){
  const relevant=ip.filter(semesterMatches).filter(matches);
  const named=relevant.filter(r=>valid(r.org));
  select('organization',named.map(r=>r.org),'All Organizations',named,'org',true);
 }
 if(page==='mou'){
  const relevant=mou.filter(semesterMatches).filter(matches);
  select('country',relevant.map(r=>r.Country),'All Countries',relevant,'Country');
  const countries=relevant.filter(r=>$('country').value==='All'||r.Country===$('country').value);
  select('moutype',countries.map(r=>r['MoU Type']),'All MoU Types',countries,'MoU Type');
 }
}
function tally(rows,key,weighted=false){const m=new Map();for(const r of rows){const k=clean(r[key]);if(!valid(k))continue;m.set(k,(m.get(k)||0)+(weighted?num(r.number):1));}return [...m].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));}
function metric(label,value,emoji='📊',caption='Current filtered selection'){const node=document.createElement('div');node.className='metric';const icon=document.createElement('span');icon.className='metric-icon';icon.textContent=emoji;const small=document.createElement('small');small.textContent=label;const strong=document.createElement('strong');strong.textContent=typeof value==='number'?fmt(value):value;const foot=document.createElement('span');foot.className='metric-caption';foot.textContent=caption;const index=document.createElement('span');index.className='metric-index';index.textContent='METRIC '+String($('metrics').children.length+1).padStart(2,'0');node.append(index,icon,small,strong,foot);$('metrics').append(node);}
function graph(title,entries,kind='horizontal',wide=false){const panel=document.createElement('article');panel.className='panel'+(wide?' wide':'');const heading=document.createElement('h3');heading.textContent=title;panel.append(heading);$('charts').append(panel);if(!entries.length)return;const wrap=document.createElement('div');wrap.className='graph-wrap';if(kind!=='doughnut')wrap.style.height=Math.max(310,entries.length*34+90)+'px';const canvas=document.createElement('canvas');wrap.append(canvas);panel.append(wrap);const labels=entries.map(x=>x[0]),values=entries.map(x=>x[1]);const colors=['#2667d7','#13a78d','#805bd3','#f0a64b','#3aa0d4','#e4677c','#4366a5','#7fbd70'];const pie=kind==='doughnut',horizontal=kind==='horizontal';charts.push(new Chart(canvas,{type:pie?'doughnut':'bar',data:{labels,datasets:[{data:values,backgroundColor:entries.map((_,i)=>colors[i%colors.length]),borderRadius:pie?0:5,maxBarThickness:38}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:horizontal?'y':'x',plugins:{legend:{display:pie,position:'bottom'},tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${fmt(ctx.raw)}`}}},scales:pie?{}:horizontal?{x:{beginAtZero:true,ticks:{precision:0}},y:{ticks:{autoSkip:false}}}:{x:{ticks:{autoSkip:false}},y:{beginAtZero:true,ticks:{precision:0}}}}}));}

function mouInsightPanel(title, subtitle, entries, mode) {
 const panel=document.createElement('article');
 panel.className='panel mou-insight-panel mou-insight-'+mode;
 const head=document.createElement('div');head.className='mou-insight-heading';
 const copy=document.createElement('div');
 const h=document.createElement('h3');h.textContent=title;
 const p=document.createElement('p');p.textContent=subtitle;
 copy.append(h,p);
 const total=entries.reduce((sum,item)=>sum+item[1],0);
 const badge=document.createElement('span');badge.className='mou-insight-count';
 badge.textContent=fmt(total)+' records';
 head.append(copy,badge);panel.append(head);
 const body=document.createElement('div');body.className='mou-insight-body';
 const palette=['#087f68','#2d64ce','#7854c4','#e79337','#168eac','#d85c72'];
 if(!entries.length) {
  const empty=document.createElement('p');empty.className='mou-insight-empty';empty.textContent='No matching MoU records';
  body.append(empty);panel.append(body);$('charts').append(panel);return;
 }
 if(mode==='types'){
  const summary=document.createElement('div');summary.className='mou-type-summary';
  const value=document.createElement('strong');value.textContent=fmt(total);
  const label=document.createElement('span');label.textContent='Total MoU records';
  summary.append(value,label);body.append(summary);
 }
 const max=Math.max(1,...entries.map(e=>e[1]));
 entries.forEach(([label,value],i)=>{
  const row=document.createElement('div');row.className='mou-insight-row';
  const top=document.createElement('div');top.className='mou-insight-rowtop';
  const l=document.createElement('span');l.className='mou-insight-name';
  const dot=document.createElement('i');dot.className='mou-insight-dot';dot.style.backgroundColor=palette[i%palette.length];
  const name=document.createElement('span');name.textContent=label;
  l.append(dot,name);
  const val=document.createElement('strong');val.className='mou-insight-value';val.textContent=fmt(value);
  top.append(l,val);
  const track=document.createElement('div');track.className='mou-insight-track';
  const fill=document.createElement('div');fill.className='mou-insight-fill';
  fill.style.width=(value/max*100)+'%';fill.style.backgroundColor=palette[i%palette.length];
  track.append(fill);
  row.append(top,track);body.append(row);
 });
 panel.append(body);$('charts').append(panel);
}

function topEntries(entries,limit=15){return entries.slice(0,limit);}
function organizationGraph(entries,selectedOrg){
 const count=Number(organizationChartLimit);
 const visible=selectedOrg?entries:entries.slice(0,count);
 const title=selectedOrg?'Internship / Placement — Selected Organization':
    `Internship / Placement — Top ${organizationChartLimit} Organizations`;
 basicBarPanel(title,'Top partner organizations ranked by reported student participation',visible);
 const panel=$('charts').lastElementChild;
 const heading=panel.querySelector('h3');
 const controls=document.createElement('div');controls.className='org-chart-controls';
 const label=document.createElement('label');label.textContent='Organizations to display';label.htmlFor='org-chart-limit';
 const selectEl=document.createElement('select');selectEl.id='org-chart-limit';selectEl.setAttribute('aria-label','Organizations to display');
 for(const [value,name] of [['10','Top 10'],['15','Top 15'],['20','Top 20']]){
   const option=new Option(name,value);selectEl.add(option);
 }
 selectEl.value=organizationChartLimit;
 selectEl.disabled=selectedOrg;
 selectEl.addEventListener('change',()=>{organizationChartLimit=selectEl.value;draw();});
 controls.append(label,selectEl);
 const hint=document.createElement('span');hint.className='org-chart-hint';
 hint.textContent=selectedOrg?'Filtered to the selected organization':
   `${visible.length} of ${entries.length} organizations shown`;
 controls.append(hint);
 heading.after(controls);

}
function drawBasic(){
 const rows=basic.filter(semesterMatches).filter(matches);
 const metrics=[
  ['Submission Records',rows.length,'Registered reports in selection','📑'],
  ['Campuses',uniq(rows.map(r=>r.Campus)).length,'Campuses represented','🏛️'],
  ['Faculties',rosterFacultyCount(),'Faculties in reporting roster','🎓'],
  ['Departments',rosterDepartmentCount(),'Departments in reporting roster','🏢']
 ];
 const target=$('metrics');target.replaceChildren();
 for(const [label,count,caption,emoji] of metrics){
  const el=document.createElement('div');el.className='metric';
  const icon=document.createElement('span');icon.className='metric-icon';icon.textContent=emoji;
  const small=document.createElement('small');small.textContent=label;
  const strong=document.createElement('strong');strong.textContent=fmt(count);
  const sub=document.createElement('span');sub.className='metric-caption';sub.textContent=caption;
  el.append(icon,small,strong,sub);target.append(el);
 }
 const f=tally(rows,'Faculty'),d=tally(rows,'Department');
 for(const name of coverageFaculties())if(($('faculty').value==='All'||$('faculty').value===name)&&!f.some(v=>v[0]===name))f.push([name,0]);
 f.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
 for(const name of coverageDepartments($('faculty').value))if(($('department').value==='All'||$('department').value===name)&&!d.some(v=>v[0]===name))d.push([name,0]);
 d.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
 basicBarPanel('Submissions by Faculty','Participation by faculty · number of submitted reports',f);
 basicBarPanel('Submissions by Department','Reporting coverage across departments',d);
 // Submission Directory has been removed; only KPI cards and charts are displayed.
}
function basicBarPanel(title,subtitle,entries){
 const panel=document.createElement('article');panel.className='panel basic-ranking-panel';
 if(title.includes('Department'))panel.classList.add('basic-department-panel');
 const head=document.createElement('div');head.className='ranking-heading';
 const headingGroup=document.createElement('div');
 const h=document.createElement('h3');h.textContent=title;
 const sub=document.createElement('p');sub.className='chart-subtitle';sub.textContent=subtitle;
 headingGroup.append(h,sub);
 const pill=document.createElement('span');pill.className='ranking-count';pill.textContent=entries.length+' '+(entries.length===1?'group':'groups');
 head.append(headingGroup,pill);panel.append(head);
 const rows=document.createElement('div');rows.className='ranking-rows';
 if(!entries.length){rows.textContent='No matching submissions';panel.append(rows);$('charts').append(panel);return;}
 const largest=Math.max(...entries.map(v=>v[1]),1);
 const colors=['#0b8f76','#2563be','#7660c8','#e69c3b','#1684a2','#285399'];
 const displayEntries=entries;
 displayEntries.forEach(([name,value],i)=>{
  const item=document.createElement('div');item.className='ranking-item';
  const number=document.createElement('span');number.className='ranking-number';number.textContent=String(i+1).padStart(2,'0');
  const detail=document.createElement('div');detail.className='ranking-detail';
  const line=document.createElement('div');line.className='ranking-line';
  const label=document.createElement('span');label.className='ranking-name';label.textContent=name;label.title=name;
  const count=document.createElement('strong');count.className='ranking-value';count.textContent=fmt(value);
  line.append(label,count);
  const track=document.createElement('div');track.className='ranking-track';
  const fill=document.createElement('div');fill.className='ranking-fill';fill.style.width=(value/largest*100)+'%';fill.style.background=colors[i%colors.length];
  track.append(fill);detail.append(line,track);item.append(number,detail);rows.append(item);
 });
 panel.append(rows);$('charts').append(panel);
}
function draw(){charts.forEach(c=>c.destroy());charts=[];$('metrics').replaceChildren();$('charts').replaceChildren();
 if(page==='basic'){drawBasic();return;}
 if(page==='ip'){
  let rows=ip.filter(semesterMatches).filter(matches);
  const selectedOrg=$('organization')&&$('organization').value!=='All';
  if(selectedOrg)rows=rows.filter(r=>r.org===$('organization').value);
  const orgRows=rows.filter(r=>valid(r.org));const orgTally=tally(orgRows,'org',true);
  metric('Total Internship / Placement — Including Planned & Remote',rows.reduce((a,b)=>a+b.number,0),'💼','Reported quantities including planned opportunities; not all are completed placements');
  metric('Faculties',rosterFacultyCount(),'🎓','Faculties in reporting roster');
  metric('Departments',rosterDepartmentCount(),'🏢','Departments in reporting roster');
  metric('Partner Organizations',uniq(orgRows.map(r=>r.org)).length,'🤝','Named external organizations');
  metric('Reporting Submissions',new Set(rows.map(r=>r.ref)).size,'📑','Source reports represented');
  organizationGraph(orgTally,selectedOrg);
  const ipFac=tally(rows,'Faculty',true);for(const name of coverageFaculties())if(($('faculty').value==='All'||$('faculty').value===name)&&!ipFac.some(v=>v[0]===name))ipFac.push([name,0]);
  ipFac.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
  basicBarPanel('Student Participation by Faculty','Reported students by faculty',ipFac);
  const deptEntries=tally(rows,'Department',true);
  for(const name of coverageDepartments($('faculty').value))if(($('department').value==='All'||$('department').value===name)&&!deptEntries.some(item=>item[0]===name))deptEntries.push([name,0]);
  deptEntries.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
  basicBarPanel('Student Participation by Department','Reported students by department',deptEntries);
  return;
 }
 let rows=mou.filter(semesterMatches).filter(matches);
 if($('country').value!=='All')rows=rows.filter(r=>r.Country===$('country').value);
 if($('moutype').value!=='All')rows=rows.filter(r=>r['MoU Type']===$('moutype').value);
 metric('MoU Records',rows.length,'📄','Records in selection');
 metric('Countries',uniq(rows.map(r=>r.Country)).length,'🌍','Countries represented');
 metric('Partner Institutions',uniq(rows.map(r=>r['Partner Institution'])).length,'🤝','External institutions');
 metric('Faculties',rosterFacultyCount(),'🎓','Faculties in reporting roster');
 metric('Departments',rosterDepartmentCount(),'🏢','Departments in reporting roster');
 metric('Reporting Submissions',new Set(rows.map(r=>r['Submission Ref'])).size,'📑','Source reports represented');
 mouInsightPanel('MoU Partnerships by Country','Geographic distribution of institutional collaborations',tally(rows,'Country'),'countries');
 mouInsightPanel('Collaboration Type Breakdown','Academic, industry and other partnership types',tally(rows,'MoU Type'),'types');
 mouInsightPanel('MoU Progress & Status','Current completion and processing stages',tally(rows,'Status'),'status');
 basicBarPanel('Partner Institutions','Collaborating institutions ranked by recorded MoUs',tally(rows,'Partner Institution'));
 const mouFac=tally(rows,'Faculty');for(const name of coverageFaculties())if(($('faculty').value==='All'||$('faculty').value===name)&&!mouFac.some(v=>v[0]===name))mouFac.push([name,0]);
 mouFac.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
 basicBarPanel('MoUs by Faculty','Collaboration activity by faculty',mouFac);
 const mouDept=tally(rows,'Department');
 for(const name of coverageDepartments($('faculty').value))if(($('department').value==='All'||$('department').value===name)&&!mouDept.some(item=>item[0]===name))mouDept.push([name,0]);
 mouDept.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
 basicBarPanel('MoUs by Department','Collaboration activity by department',mouDept);
}
function reset(){for(const id of ['semester','faculty','department','organization','country','moutype'])if($(id))$(id).value='All';updateFilters();draw();}
async function boot(){
 try{
  if(typeof XLSX==='undefined')throw Error('Excel library did not load. Check internet access or local library.');
  if(typeof Chart==='undefined')throw Error('Chart.js did not load. Check internet access or local library.');
  const source=await workbook('data/64_Submissions_Source.xlsx');
  roster=sheet(source,'All Submissions').filter(r=>valid(r['Ref #']));
  submissionByRef=new Map(roster.map(r=>[clean(r['Ref #']),r]));
  departmentRoster=roster.filter(r=>valid(r.Department)).map(r=>({Department:clean(r.Department),Faculty:clean(r.Faculty)}));
  facultyRoster=uniq(roster.map(r=>r.Faculty));
   updateSemesterLabels();
  const tagSemester=r=>{
  const ref=clean(r.ref||r['Submission Ref']||r['Ref #']);
  const master=submissionByRef.get(ref);
  // PRPSTFPT34's original semester field is Spring 2026 although submitted in July.
  // Use the department's explicit semester selection for this report in every section.
  const sourcePeriod=master ? semesterForSubmission(master) : semesterOf(r['Reporting Month']);
  return {...r,semester:sourcePeriod};
};
  if(page==='basic'){
   basic=roster.map(tagSemester);
   if(!basic.length)throw Error('No records found in the All Submissions sheet');
  }else if(page==='ip'){
   const b=await workbook('data/UOL_Separated_All_Activity_Data_Cleaned.xlsx');
   // Keep all source categories intact in Excel; merge only for dashboard reporting.
   const internship=sheet(b,'Internship - Placement');
   const planned=sheet(b,'Planned - Section 4'); // Structured sheet with correct headers; planned quantities are deduplicated in parseIp.
   const remote=sheet(b,'Remote Activities');
   ip=parseIp([...internship, ...planned, ...remote]).filter(r=>submissionByRef.has(r.ref)).map(tagSemester);
   if(!ip.length)throw Error('No records found in Internship / Placement, Planned or Remote sheets');
  }else if(page==='mou'){
   const c=await workbook('data/UOL_MoU_Country_Previous.xlsx');
   mou=sheet(c,'MoU Country (No Nulls)').filter(r=>valid(r.Country)&&submissionByRef.has(clean(r['Submission Ref']))).map(tagSemester);
   if(!mou.length)throw Error('No valid country records found in MoU country sheet');
  }
  $('connection').textContent=`Excel Connected · ${page==='basic'?basic.length:page==='ip'?ip.length:mou.length} records`;
  updateFilters();
  for(const id of ['semester','faculty','department','organization','country','moutype'])$(id)?.addEventListener('change',()=>{
   if(id==='faculty'||id==='department'||id==='country'||id==='semester')updateFilters(id);
   draw();
  });
  $('reset')?.addEventListener('click',reset);
  if(page==='mou' && $('generateMouPdf')){
   const pdfButton=$('generateMouPdf');pdfButton.disabled=false;
   pdfButton.addEventListener('click',async()=>{
    const rows=mou.filter(semesterMatches).filter(matches).filter(r=>($('country').value==='All'||r.Country===$('country').value)&&($('moutype').value==='All'||r['MoU Type']===$('moutype').value));
    const filters={semester:selectedSemester(),faculty:$('faculty').value,department:$('department').value,country:$('country').value,type:$('moutype').value};
    pdfButton.disabled=true;pdfButton.textContent='Preparing PDF…';
    try{await window.generateMouPdf({rows,filters});}catch(error){console.error('MoU PDF export failed:',error);alert('Could not generate PDF: '+error.message);}
    finally{pdfButton.disabled=false;pdfButton.textContent='⬇ Generate PDF Report';}
   });
  }
  draw();
 }catch(e){
  console.error('Section dashboard load failed:',e);
  $('connection').textContent='Data load error: '+e.message;
  $('connection').style.background='#ffe4e6';
  const target=$('charts');
  if(target){const error=document.createElement('p');error.style.color='#b91c1c';error.textContent=e.message;target.append(error);}
 }
}
// Student Engagement sidebar uses the same cleaned Excel workbook as Section 2.
// It is independent of each page's filters and never blocks the primary report.
async function loadSidebarEngagement(){
 const target=$('sideStudents');if(!target)return;
 try{
  const b=await workbook('data/UOL_Separated_All_Activity_Data_Cleaned.xlsx');
  const rows=page==='ip'&&ip.length?ip:(()=>{const main=sheet(b,'Internship - Placement'),planned=sheet(b,'Planned - Section 4');return parseIp([...main,...planned]).filter(r=>submissionByRef.has(r.ref));})();
  const total=rows.reduce((sum,row)=>sum+num(row.number),0);
  target.textContent=fmt(total);
 }catch(err){console.warn('Student engagement sidebar could not load:',err);target.textContent='—';}
}
const originalBoot=boot;
boot=async function(){await originalBoot();await loadSidebarEngagement();};
document.querySelector(`[data-page="${page}"]`)?.classList.add('active');boot();
})();