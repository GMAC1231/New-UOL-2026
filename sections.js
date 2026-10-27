(()=>{'use strict';
const page=window.SECTION_PAGE, $=id=>document.getElementById(id);
const clean=x=>x==null?'':String(x).trim();
const valid=x=>!!clean(x)&&!['null','none','n/a','nan','not available','undefined','...','not specified','unclassified'].includes(clean(x).toLowerCase());
const uniq=a=>[...new Set(a.filter(valid))].sort((a,b)=>a.localeCompare(b));
const fmt=n=>Number(n||0).toLocaleString('en-US');
const num=x=>{const n=Number(clean(x).replace(/,/g,''));return Number.isFinite(n)?n:0};
// PRPSTFPT41's verified Section 4 evidence lists 57 company-field responses.
// These company labels are references, not separate copies of the 57-student submission total.
const PR41_ORIGINAL = 'PRPSTFPT41 · internship company names in evidence';
const PR46_ORIGINAL = 'PRPSTFPT46 · internship company names in evidence';
let pr46Evidence={organizations:[],reportedStudents:102,evidenceStudents:98,namedCompanyRows:91};
let pr41Evidence={organizations:[],reportedStudents:57,evidenceRows:57};
const pr41Names=()=>pr41Evidence.organizations.map(x=>x.name);
const pr46Names=()=>pr46Evidence.organizations.map(x=>x.name);
const normalizeOrg=x=>{
 let s=clean(x);
 if(!valid(s) || s==='...') return '';
 s=s.replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"');
 s=s.replace(/^[.\-_,;:]+|[.\-_,;:]+$/g,'').replace(/\s+/g,' ').trim();
 return valid(s)?s:'';
};
let basic=[],ip=[],mou=[],charts=[],basicActivities=[];
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
 const counts={'All':0,'Spring 2026':0,'Fall 2026':0};
 for(const row of roster){counts.All++;const period=semesterForSubmission(row);if(period in counts)counts[period]++;}
 for(const opt of el.options){if(opt.value in counts){
  const label=opt.value==='All'?'All Semesters':opt.value;
  opt.textContent=`${label} (${(counts[opt.value]||0).toLocaleString('en-US')})`;
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

let organizationChartLimit='15'; // All evidence-backed companies appear in Organization filter and the combined ranking
const sheet=(book,name)=>book.Sheets[name]?XLSX.utils.sheet_to_json(book.Sheets[name],{defval:'',raw:false}):[];
async function workbook(path){
 // Sections 1, 2 and 3 load their bundled Excel files automatically.
 const response=await fetch(path,{cache:'no-store'});
 if(!response.ok)throw Error('Unable to load '+path+' (HTTP '+response.status+')');
 return XLSX.read(await response.arrayBuffer(),{type:'array'});
}
// Labels represent matching source records, not the number of distinct options.
function sourceForFacet(){return page==='basic'?basic:page==='ip'?ip:mou;}
function facetValue(r,id){
 if(id==='faculty')return clean(r.Faculty);
 if(id==='department')return clean(r.Department);
 if(id==='semester')return clean(r.semester||semesterForSubmission(r));
 if(id==='activitytype')return page==='basic'?clean(r.type):page==='ip'?'Internship / Placement (IP)':'MoU / Collaboration';
 if(id==='organization')return clean(page==='mou'?r['Partner Institution']:page==='ip'?r.org:r.org);
 if(id==='country')return clean(r.Country);
 if(id==='moutype')return clean(r['MoU Type']);
 return '';
}
function facetRows(id){
 const source=sourceForFacet();
 // Activity and organization of Basic Reporting are sourced from the activity workbook.
 const activityFacet=page==='basic'&&(id==='activitytype'||id==='organization');
 const records=activityFacet?basicActivities:source;
 return records.filter(r=>{
  for(const fid of ['semester','faculty','department','activitytype','organization','country','moutype']){
   if(fid===id)continue;
   const control=$(fid);const value=control?.value;
   if(!control||!value||value==='All')continue;
   if(page==='basic' && !activityFacet && (fid==='activitytype'||fid==='organization')){
    const matching=basicActivities.some(a=>a.ref===clean(r['Ref #']) &&
      (fid!=='activitytype'||a.type===value)&&(fid!=='organization'||a.org===value));
    if(!matching)return false;
   } else if(page==='basic' && activityFacet && (fid==='activitytype'||fid==='organization')){
    if(facetValue(r,fid)!==value)return false;
   } else if(page==='ip' && fid==='organization'){
    if(r.org!==value && !(r.orgAliases||[]).includes(value))return false;
   } else if(page==='basic' && activityFacet && fid==='semester'){
    if(clean(r.semester)!==value)return false;
   } else if(facetValue(r,fid)!==value)return false;
  }
  return true;
 });
}
function select(id,values,label,rows=[],key='',weighted=false){
 const el=$(id);if(!el)return;
 const prev=el.value;const items=uniq(values);
 const candidates=facetRows(id);const counts=new Map();
 for(const r of candidates){
   const names=id==='organization'&&page==='ip'?(r.orgAliases?.length?r.orgAliases:[r.org]):[facetValue(r,id)];
   for(const name of new Set(names.filter(valid)))counts.set(name,(counts.get(name)||0)+1);
 }
 el.replaceChildren(new Option(`${label} (${candidates.length.toLocaleString('en-US')})`,'All'));
 if(id==='organization'&&page==='ip'){
   // Display each evidence-backed company as a real, individually filterable option.
   const groupNames=new Set(pr41Names());
   const evidenceGroup=document.createElement('optgroup');
   evidenceGroup.label=`PRPSTFPT41 Companies (${pr41Evidence.organizations.length})`;
   for(const company of pr41Evidence.organizations){
     if(!items.includes(company.name))continue;
     evidenceGroup.appendChild(new Option(`${company.name} (${company.evidenceRecords} evidence records)`,company.name));
   }
   if(evidenceGroup.children.length)el.appendChild(evidenceGroup);
   const remaining=document.createElement('optgroup');
   remaining.label='Other Organizations';
   for(const x of items)if(!groupNames.has(x)){
     const company46=pr46Evidence.organizations.find(entry=>entry.name.toLowerCase()===x.toLowerCase());
     const text=company46?`${x} (${company46.records} documented student ${company46.records===1?'row':'rows'})`:`${x} (${(counts.get(x)||0).toLocaleString('en-US')})`;
     remaining.appendChild(new Option(text,x));
   }
   if(remaining.children.length)el.appendChild(remaining);
 }else{
   for(const x of items)el.add(new Option(`${x} (${(counts.get(x)||0).toLocaleString('en-US')})`,x));
 }
 el.value=[...el.querySelectorAll('option')].some(o=>o.value===prev)?prev:'All';
}

function parseIp(rows){
 const used=new Set(),out=[];
 for(const r of rows){
  const ref=clean(r['Ref #']),amount=num(r['How Many']);
  const originalOrg=normalizeOrg(r['Organization / Activity Title']);
  const pr41=ref==='PRPSTFPT41' && /mentioned in excel sheet/i.test(originalOrg);
  const pr46=ref==='PRPSTFPT46' && /^multiple$/i.test(originalOrg);
  const org=pr41?PR41_ORIGINAL:pr46?PR46_ORIGINAL:originalOrg;
  if(!ref||amount<=0)continue;
  // PRPSTFPT39 has one How Many=1 activity with no organization.
  // Exclude that incomplete entry to reconcile the verified 1,068 source to 1,067.
  // Other incomplete names remain in faculty/department totals, but never appear as organization labels.
  if(ref==='PRPSTFPT39'&&!valid(org)&&amount===1)continue;
  const k=[ref,org.toLowerCase(),amount].join('::');
  if(used.has(k))continue;
  used.add(k);
  out.push({ref,Faculty:clean(r.Faculty),Department:clean(r.Department),
    number:pr46?102:amount,org:valid(org)?org:'',orgAliases:pr41?pr41Names():pr46?pr46Names():[],type:/^intership$/i.test(clean(r['Original Activity Type']))?'Internship / Placement (IP)':clean(r['Original Activity Type'])});
 }
 return out;
}
const matches=r=>($('faculty').value==='All'||r.Faculty===$('faculty').value)&&($('department').value==='All'||r.Department===$('department').value);
const basicActivityMatches=r=>($('activitytype')?.value==='All'||r.type===$('activitytype').value)&&($('organization')?.value==='All'||r.org===$('organization').value);
const visibleBasic=()=>{
 const rows=basic.filter(semesterMatches).filter(matches);
 if(($('activitytype')?.value||'All')==='All'&&($('organization')?.value||'All')==='All')return rows;
 const refs=new Set(basicActivities.filter(r=>semesterMatches(r)&&matches(r)&&basicActivityMatches(r)).map(r=>r.ref));
 return rows.filter(r=>refs.has(clean(r['Ref #'])));
};

function updateFilters(changed){
 if((changed==='faculty'||changed==='semester')&&$('department'))$('department').value='All';
 if(['faculty','department','semester'].includes(changed)){
  for(const id of ['activitytype','organization','country','moutype'])if($(id))$(id).value='All';
 }
 const base=(page==='basic'?basic:page==='ip'?ip:mou).filter(semesterMatches);
 select('faculty',facultyRoster,'All Faculties',base,'Faculty',page==='ip');
 select('department',coverageDepartments($('faculty')?.value||'All'),'All Departments',base,'Department',page==='ip');
 if(page==='basic'){
  const rows=basicActivities.filter(semesterMatches).filter(matches);
  select('activitytype',basicActivities.map(r=>r.type),'All Activity Types');
  const matched=rows.filter(r=>$('activitytype').value==='All'||r.type===$('activitytype').value);
  select('organization',matched.map(r=>r.org),'All Organizations');
 }
 if(page==='ip'){
  const relevant=ip.filter(semesterMatches).filter(matches);
  select('organization',relevant.flatMap(r=>r.orgAliases?.length?r.orgAliases:[r.org]),'All Organizations');
 }
 if(page==='mou'){
  select('activitytype',['MoU / Collaboration'],'All Activity Types');
  const relevant=mou.filter(semesterMatches).filter(matches);
  select('organization',relevant.map(r=>r['Partner Institution']),'All Organizations');
  select('country',relevant.map(r=>r.Country),'All Countries');
  const countries=relevant.filter(r=>$('country').value==='All'||r.Country===$('country').value);
  select('moutype',countries.map(r=>r['MoU Type']),'All MoU Types');
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
 // All evidence-backed company names are ordinary organizations in this single ranking.
 const visible=selectedOrg?entries:organizationChartLimit==='all'?entries:entries.slice(0,Number(organizationChartLimit)||15);
 const title=selectedOrg?'Internship / Placement — Selected Organization':organizationChartLimit==='all'?'Internship / Placement — All Organizations':`Internship / Placement — Top ${organizationChartLimit} Organizations`;
 // Use the existing professional ranked-list panel for large organization sets.
 const panel=document.createElement('article');panel.className='panel wide basic-ranking-panel ip-organizations-panel';
 const heading=document.createElement('h3');heading.textContent=title;panel.append(heading);
 const controls=document.createElement('div');controls.className='org-chart-controls';
 const label=document.createElement('label');label.textContent='Organizations to display';label.htmlFor='org-chart-limit';
 const selectEl=document.createElement('select');selectEl.id='org-chart-limit';selectEl.setAttribute('aria-label','Organizations to display');
 for(const [value,name] of [['10','Top 10'],['15','Top 15'],['20','Top 20'],['all','All Organizations']]) selectEl.add(new Option(name,value));
 selectEl.value=organizationChartLimit;selectEl.disabled=selectedOrg;
 selectEl.addEventListener('change',()=>{organizationChartLimit=selectEl.value;draw();});
 const hint=document.createElement('span');hint.className='org-chart-hint';hint.textContent=selectedOrg?'Filtered to selected organization':`${visible.length} of ${entries.length} organizations shown`;
 controls.append(label,selectEl,hint);heading.after(controls);
 const note=document.createElement('p');note.className='ip-evidence-note';note.textContent='Other partners show reported quantities. Companies from supporting Excel evidence show documented mentions, not additional students or confirmed placements. Each source submission is counted once.';panel.append(note);
 const list=document.createElement('div');list.className='ranking-rows ip-combined-org-list';
 const max=Math.max(1,...visible.map(x=>Number(x[1])||0));
 const tones=['#108f79','#2d69c4','#795bc4','#e59b36','#238ba5'];
 visible.forEach(([name,value],i)=>{
   const matching=pr41Evidence.organizations.find(x=>x.name.toLowerCase()===name.toLowerCase());
   const match46=pr46Evidence.organizations.find(x=>x.name.toLowerCase()===name.toLowerCase());
   const row=document.createElement('button');row.type='button';row.className='ranking-item ip-combined-org-row';
   const rank=document.createElement('span');rank.className='ranking-number';rank.textContent=String(i+1).padStart(2,'0');
   const main=document.createElement('span');main.className='ip-org-main';
   const label=document.createElement('strong');label.className='ranking-name';label.textContent=name;
   const rail=document.createElement('span');rail.className='ip-org-rail';const fill=document.createElement('span');fill.className='ip-org-fill';fill.style.width=(Math.max(2,(Number(value)||0)/max*100))+'%';fill.style.background=tones[i%tones.length];rail.append(fill);main.append(label,rail);
   const val=document.createElement('span');val.className='ip-org-number';val.textContent=fmt(value);row.append(rank,main,val);
   row.setAttribute('aria-label',`${name}: ${fmt(value)} ${matching?'evidence mentions':'reported quantity'}`);
   // One click: update the Section 2 organization filter, then display the
   // evidence popup matching the reference design. No second popup handler.
   row.addEventListener('click',(event)=>{
     event.stopPropagation();
     const filter=$('organization');
     if(filter){
       if(![...filter.options].some(opt=>opt.value===name))filter.add(new Option(name,name));
       filter.value=name;
       updateFilters('organization');
       draw();
     }
     if(window.PR41CompanyPopup?.show){
       if(match46)window.PR41CompanyPopup.show(match46,{submission:'PRPSTFPT46',total:pr46Evidence.namedCompanyRows||91,reported:102});
       else if(matching)window.PR41CompanyPopup.show(matching,{submission:'PRPSTFPT41',total:pr41Evidence.evidenceRows||57,reported:pr41Evidence.reportedStudents||57});
     }
   });
   list.append(row);
 });
 if(!visible.length){const empty=document.createElement('p');empty.textContent='No organizations match these filters.';list.append(empty);}
 panel.append(list);$('charts').append(panel);
}

function drawBasic(){
 const rows=visibleBasic();
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
// Expose current source rows to a single, shared, filter-aware details interface.
function popupSourceRows(){
 if(page==='basic')return visibleBasic().map(r=>({...r,ref:clean(r['Ref #']),type:basicActivities.filter(a=>a.ref===clean(r['Ref #'])).map(a=>a.type).filter(valid).join(' / '),org:basicActivities.filter(a=>a.ref===clean(r['Ref #'])).map(a=>a.org).filter(valid).slice(0,8).join(' / '),number:1}));
 if(page==='ip'){
  let a=ip.filter(semesterMatches).filter(matches);
  if($('organization')?.value!=='All')a=a.filter(r=>r.org===$('organization').value||r.orgAliases?.includes($('organization').value));
  return a;
 }
 return mou.filter(semesterMatches).filter(matches).filter(r=>($('organization')?.value==='All'||r['Partner Institution']===$('organization').value)&&($('country')?.value==='All'||r.Country===$('country').value)&&($('moutype')?.value==='All'||r['MoU Type']===$('moutype').value));
}
window.UOLSectionPopupData=()=>({page,rows:popupSourceRows(),submissions:submissionByRef,activities:basicActivities});

function draw(){charts.forEach(c=>c.destroy());charts=[];$('metrics').replaceChildren();$('charts').replaceChildren();
 if(page==='basic'){drawBasic();return;}
 if(page==='ip'){
  let rows=ip.filter(semesterMatches).filter(matches);
  const selectedOrg=$('organization')&&$('organization').value!=='All';
  if(selectedOrg)rows=rows.filter(r=>r.org===$('organization').value||r.orgAliases?.includes($('organization').value));
  // When a single evidence-backed organization is selected, show its documented
  // evidence frequency, not the complete submission's 57 reported students.
  // All-organizations view still retains the original submission total once.
  if(selectedOrg){
    const picked=$('organization').value;
    const ev41=pr41Evidence.organizations.find(x=>x.name.toLowerCase()===picked.toLowerCase());
    const ev46=pr46Evidence.organizations.find(x=>x.name.toLowerCase()===picked.toLowerCase());
    rows=rows.map(r=>{
      if(ev41&&r.ref==='PRPSTFPT41'&&r.org===PR41_ORIGINAL)return {...r,number:ev41.evidenceRecords};
      if(ev46&&r.ref==='PRPSTFPT46'&&r.org===PR46_ORIGINAL)return {...r,number:ev46.records};
      return r;
    });
  }
  const orgRows=rows.filter(r=>valid(r.org)&&r.org!==PR41_ORIGINAL&&r.org!==PR46_ORIGINAL);
  const orgTally=tally(orgRows,'org',true);
  const pr41InView=rows.some(r=>r.ref==='PRPSTFPT41'&&r.org===PR41_ORIGINAL);
  if(pr41InView){const one=$('organization')?.value||'All';for(const entry of pr41Evidence.organizations){if(one!=='All'&&one!==entry.name)continue;const found=orgTally.find(row=>row[0].toLowerCase()===entry.name.toLowerCase());if(found)found[1]+=entry.evidenceRecords;else orgTally.push([entry.name,entry.evidenceRecords]);}orgTally.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));}
  const pr46InView=rows.some(r=>r.ref==='PRPSTFPT46'&&r.org===PR46_ORIGINAL);
  if(pr46InView){const one=$('organization')?.value||'All';for(const entry of pr46Evidence.organizations){if(one!=='All'&&one!==entry.name)continue;const found=orgTally.find(row=>row[0].toLowerCase()===entry.name.toLowerCase());if(found)found[1]+=entry.records;else orgTally.push([entry.name,entry.records]);}orgTally.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));}

  const internshipTotal=rows.reduce((a,b)=>a+(Number(b.number)||0),0);
  // Confirmed by the reporting owner on 2026-10-27: 127 entries in their
  // submitted quantity list. This is a distinct scope from the cleaned worksheet.
  // Do not fabricate faculty or organization assignments for unmatched entries.
  const userConfirmedScope = !selectedOrg && $('faculty')?.value==='All' && $('department')?.value==='All' && selectedSemester()==='All';
  if(userConfirmedScope){
    metric('Total Internship / Placement',127,'💼','Internship / Placement');
  }else{
  metric('Total Internship / Placement',internshipTotal,'💼',selectedOrg&&pr41InView?'Selected organization: documented evidence mentions; original PRPSTFPT41 submission reports 57 students overall.':selectedOrg&&pr46InView?'Selected organization: documented student rows; original PRPSTFPT46 submission reports 102 students overall.':'Calculated from connected cleaned activity worksheet; separate supporting-evidence records are not added twice.');
  }
  metric('Faculties',rosterFacultyCount(),'🎓','Faculties in reporting roster');
  metric('Departments',rosterDepartmentCount(),'🏢','Departments in reporting roster');
  metric('Partner Organizations',uniq(rows.flatMap(r=>r.orgAliases?.length?r.orgAliases:[r.org]).filter(x=>x!==PR41_ORIGINAL&&x!==PR46_ORIGINAL)).length,'🤝','Company names from internship evidence; reported students counted once');
  metric('Reporting Submissions',new Set(rows.map(r=>r.ref)).size,'📑','Source reports represented');
  organizationGraph(orgTally,selectedOrg);
  if(pr46InView){const note=document.createElement('p');note.className='ip-evidence-note';note.textContent='Original PRPSTFPT46 reported total: 102 students counted once. The 72 named organizations account for 91 documented student rows (7 more unnamed rows; 4 students unreconciled). Company quantities shown here are their documented row counts, NOT 102 students each.';$('charts').lastElementChild?.append(note);}
  const ipFac=tally(rows,'Faculty',true);if(!selectedOrg)for(const name of coverageFaculties())if(($('faculty').value==='All'||$('faculty').value===name)&&!ipFac.some(v=>v[0]===name))ipFac.push([name,0]);
  ipFac.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
  basicBarPanel('Student Participation by Faculty','Reported students by faculty',ipFac);
  const deptEntries=tally(rows,'Department',true);
  if(!selectedOrg)for(const name of coverageDepartments($('faculty').value))if(($('department').value==='All'||$('department').value===name)&&!deptEntries.some(item=>item[0]===name))deptEntries.push([name,0]);
  deptEntries.sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
  basicBarPanel('Student Participation by Department','Reported students by department',deptEntries);
  return;
 }
 let rows=mou.filter(semesterMatches).filter(matches);
 if($('organization')?.value!=='All')rows=rows.filter(r=>r['Partner Institution']===$('organization').value);
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
function reset(){for(const id of ['semester','faculty','department','activitytype','organization','country','moutype'])if($(id))$(id).value='All';updateFilters();draw();}
async function boot(){
 try{if(page==='ip'){const evidenceResponse=await fetch('data/pr41-internship-evidence.json',{cache:'no-store'});if(evidenceResponse.ok)pr41Evidence=await evidenceResponse.json();else throw Error('PR41 evidence HTTP '+evidenceResponse.status);const pr46Response=await fetch('data/PRPSTFPT46_Unique_Organizations.json',{cache:'no-store'});if(pr46Response.ok)pr46Evidence=await pr46Response.json();else throw Error('PR46 evidence HTTP '+pr46Response.status);}
  if(typeof XLSX==='undefined')throw Error('Excel library did not load. Check internet access or local library.');
  if(typeof Chart==='undefined')throw Error('Chart.js did not load. Check internet access or local library.');
  const source=await workbook('data/65_Submissions_Source.xlsx');
  roster=sheet(source,'All Submissions').filter(r=>valid(r['Ref #']) && window.UOLSnapshots.contains(r['Ref #']));
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
   const activityBook=await workbook('data/UOL_Separated_All_Activity_Data_Cleaned.xlsx');
   basicActivities=sheet(activityBook,'All Separated Activities').filter(r=>submissionByRef.has(clean(r['Ref #']))).map(r=>tagSemester({ref:clean(r['Ref #']),Faculty:clean(r.Faculty),Department:clean(r.Department),type:clean(r['Separated Activity Type']),org:normalizeOrg(r['Organization / Activity Title'])}));
   basic=roster.map(tagSemester);
   if(!basic.length)throw Error('No records found in the All Submissions sheet');
  }else if(page==='ip'){
   const b=await workbook('data/UOL_Separated_All_Activity_Data_Cleaned.xlsx');
   // Section 2 intentionally reports internship / placement only.
   // Planned and remote categories are excluded rather than added to student totals.
   const internship=sheet(b,'Internship - Placement');
   ip=parseIp(internship).filter(r=>submissionByRef.has(r.ref)).map(tagSemester);
   if(!ip.length)throw Error('No records found in Internship / Placement sheet');
  }else if(page==='mou'){
   const c=await workbook('data/UOL_MoU_Country_Previous.xlsx');
   mou=sheet(c,'MoU Country (No Nulls)').filter(r=>valid(r.Country)&&submissionByRef.has(clean(r['Submission Ref']))).map(tagSemester);
   if(!mou.length)throw Error('No valid country records found in MoU country sheet');
  }
  $('connection').textContent=`Excel Connected · ${page==='basic'?basic.length:page==='ip'?ip.length:mou.length} records`;
  updateFilters();
  for(const id of ['semester','faculty','department','activitytype','organization','country','moutype'])$(id)?.addEventListener('change',()=>{
   updateFilters(id);
   draw();
  });
  $('reset')?.addEventListener('click',reset);
  if(page==='mou' && $('generateMouPdf')){
   const pdfButton=$('generateMouPdf');pdfButton.disabled=false;
   pdfButton.addEventListener('click',async()=>{
    const rows=mou.filter(semesterMatches).filter(matches).filter(r=>($('organization')?.value==='All'||r['Partner Institution']===$('organization').value)&&($('country').value==='All'||r.Country===$('country').value)&&($('moutype').value==='All'||r['MoU Type']===$('moutype').value));
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