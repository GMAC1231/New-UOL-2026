(()=>{'use strict';
const $=id=>document.getElementById(id), clean=x=>String(x??'').trim();
const escapeHtml=s=>clean(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const meaningful=x=>!!clean(x)&&!['null','none','nan','n/a','not available','undefined'].includes(clean(x).toLowerCase());
const cols=['Faculty','Department','Campus','Focal Person','Contact Email','Section 2: Activity Dashboard','Section 3: MoU Progress and Collaborations','Section 4: Internships & Placement Facilitation','Alumni engagement Summary','Section 6: Research & Joint Projects','Section 10: Plans for Next Month'];
const auditSections=cols.slice(5);
let rows=[],duplicates=[],missing=[];
const csv=items=>items.map(r=>r.map(x=>'"'+clean(x).replace(/"/g,'""')+'"').join(',')).join('\r\n');
function download(name,lines){const url=URL.createObjectURL(new Blob(['\uFEFF'+csv(lines)],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),2000)}
function table(head,body){return `<div class="audit-table-wrap"><table><thead><tr>${head.map(h=>`<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${body.map(r=>`<tr>${r.map(v=>`<td>${escapeHtml(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;}
function render(){
 const f=$('faculty').value, dept=$('department').value;
 const subset=rows.filter(r=>(f==='All'||clean(r.Faculty)===f)&&(dept==='All'||clean(r.Department)===dept));
 const refs=new Set(subset.map(r=>clean(r['Ref #']))), currentDup=duplicates.filter(d=>d.refs.some(ref=>refs.has(ref)));
 const missingPrimary=subset.filter(r=>!meaningful(r.Faculty)||!meaningful(r.Department));
 const blanks=cols.map(c=>[c,subset.filter(r=>!meaningful(r[c])).length]);
 $('stats').innerHTML=[['Submission records',subset.length],['Unique departments',new Set(subset.map(r=>clean(r.Department)).filter(meaningful)).size],['Repeated section descriptions',currentDup.reduce((sum,d)=>sum+d.extra,0)],['Missing faculty/department',missingPrimary.length]].map(([title,num])=>`<article class="audit-stat"><span>${escapeHtml(title)}</span><strong>${num}</strong></article>`).join('');
 $('missingTable').innerHTML=table(['Field','Blank / not provided','Filled'],blanks.map(([name,count])=>[name,count,subset.length-count]));
 $('duplicateTable').innerHTML=table(['Reporting section','Repeated submissions','Extra repeated descriptions','Review status'],currentDup.map(d=>[d.section,d.refs.join(', '),d.extra,'Review — no original records deleted']));
 $('missingRefs').innerHTML=table(['Submission ID','Faculty','Department','Focal Person'],missingPrimary.map(r=>[r['Ref #'],r.Faculty||'Missing',r.Department||'Missing',r['Focal Person']]));
 $('dupExport').onclick=()=>download('UOL_Duplicate_Descriptions_Audit.csv',[['Section','Submission IDs','Repeated descriptions','Status'],...currentDup.map(d=>[d.section,d.refs.join('; '),d.extra,'Review'])]);
 $('missingExport').onclick=()=>download('UOL_Missing_Values_Audit.csv',[['Field','Blank','Filled'],...blanks]);
}
function updateDepts(){const f=$('faculty').value;const departments=[...new Set(rows.filter(r=>f==='All'||clean(r.Faculty)===f).map(r=>clean(r.Department)).filter(meaningful))].sort();$('department').replaceChildren(new Option('All Departments','All'),...departments.map(v=>new Option(v,v)));render();}
async function init(){try{const response=await fetch('data/64_Submissions_Source.xlsx',{cache:'no-store'});if(!response.ok)throw Error('Workbook request failed');const book=XLSX.read(await response.arrayBuffer(),{type:'array'});if(!book.Sheets['All Submissions'])throw Error('All Submissions sheet missing');rows=XLSX.utils.sheet_to_json(book.Sheets['All Submissions'],{defval:'',raw:false});
 const groups=new Map();for(const r of rows){for(const section of auditSections){const value=clean(r[section]).replace(/\s+/g,' ').toLowerCase();if(!meaningful(value))continue;const key=section+'\0'+value;let g=groups.get(key);if(!g){g={section,refs:[],extra:0};groups.set(key,g);}g.refs.push(clean(r['Ref #']));}}
 duplicates=[...groups.values()].filter(g=>g.refs.length>1).map(g=>({...g,extra:g.refs.length-1}));
 $('connection').textContent=`Excel connected · ${rows.length} submissions`;const fs=[...new Set(rows.map(r=>clean(r.Faculty)).filter(meaningful))].sort();$('faculty').replaceChildren(new Option('All Faculties','All'),...fs.map(v=>new Option(v,v)));$('faculty').onchange=updateDepts;$('department').onchange=render;updateDepts();
}catch(e){$('connection').textContent='Excel unavailable';$('error').textContent=e.message;console.error(e)}}init();
})();
