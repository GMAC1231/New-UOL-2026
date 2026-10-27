(()=>{'use strict';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const get=()=>window.UOLSectionPopupData?.();
const root=document.createElement('div');root.className='uol-major-overlay';root.hidden=true;root.innerHTML='<section class="uol-major-dialog" role="dialog" aria-modal="true" aria-labelledby="uolMajorTitle"><div class="uol-major-head"><div><small>UOL · SOURCE EVIDENCE</small><h2 id="uolMajorTitle">Details</h2><p id="uolMajorSummary"></p></div><button type="button" class="uol-major-close" aria-label="Close details">×</button></div><div class="uol-major-tools"><input id="uolMajorSearch" placeholder="Search focal person, activity, department, organization…" aria-label="Search popup records"><button id="uolMajorExport" type="button">Export CSV</button></div><div class="uol-major-scroll"><table><thead><tr><th>Ref ID</th><th>Faculty</th><th>Department</th><th>Focal Person</th><th>Activity</th><th>Organization / Partner</th><th>Quantity</th><th>Semester</th></tr></thead><tbody id="uolMajorRows"></tbody></table></div><p class="uol-major-note">Source records and reported quantities are separate. Repeated source entries are shown for review, not automatically deleted.</p></section>';
document.body.append(root);
let full=[],shown=[];
const val=(r,k,ctx)=>{
 const master=ctx?.submissions?.get(String(r.ref||r['Submission Ref']||r['Ref #']||''))||{};
 switch(k){
 case 'Ref ID':return r.ref||r['Submission Ref']||r['Ref #']||'';
 case 'Faculty':return r.Faculty||'';
 case 'Department':return r.Department||'';
 case 'Focal Person':return r['Focal Person']||master['Focal Person']||master['Focal Person Name']||'';
 case 'Activity':return r.type||r['MoU Type']||r['Section 2: Activity Dashboard']||(ctx.page==='basic'?'Submission Record':'MoU / Collaboration');
 case 'Organization / Partner':return r.org||r['Partner Institution']||r['Organization / Activity Title']||'';
 case 'Quantity':return ctx.page==='basic'?'1':r.number??r['How Many']??'1';
 case 'Semester':return r.semester||'';
 }
};
const cols=['Ref ID','Faculty','Department','Focal Person','Activity','Organization / Partner','Quantity','Semester'];
function paint(){const query=document.getElementById('uolMajorSearch').value.trim().toLowerCase(),ctx=get();shown=full.filter(r=>!query||cols.some(k=>String(val(r,k,ctx)).toLowerCase().includes(query)));document.getElementById('uolMajorRows').innerHTML=shown.slice(0,600).map(r=>'<tr>'+cols.map(k=>'<td>'+esc(val(r,k,ctx)||'—')+'</td>').join('')+'</tr>').join('')||'<tr><td colspan="8">No matching source records</td></tr>';document.getElementById('uolMajorSummary').textContent=`${shown.length.toLocaleString()} matching records · ${full.length.toLocaleString()} records in this selection`}
function open(title,filter){const ctx=get();if(!ctx)return;full=ctx.rows.filter(r=>!filter||filter(r,ctx));document.getElementById('uolMajorTitle').textContent=title;document.getElementById('uolMajorSearch').value='';root.hidden=false;document.body.classList.add('uol-major-open');paint()}
function close(){root.hidden=true;document.body.classList.remove('uol-major-open')}
root.querySelector('.uol-major-close').onclick=close;root.addEventListener('click',e=>{if(e.target===root)close()});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!root.hidden)close()});document.getElementById('uolMajorSearch').addEventListener('input',paint);
document.getElementById('uolMajorExport').onclick=()=>{const ctx=get();const csv=[cols,...shown.map(r=>cols.map(k=>val(r,k,ctx)))].map(a=>a.map(v=>'"'+String(v??'').replaceAll('"','""')+'"').join(',')).join('\r\n');const u=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=u;a.download='uol_filtered_source_records.csv';a.click();setTimeout(()=>URL.revokeObjectURL(u),1000)};
// Delegate clicks so dynamically redrawn charts/cards keep their actions.
document.addEventListener('click',e=>{if(!root.hidden)return;const metric=e.target.closest('#metrics .metric');if(metric){open((metric.querySelector('small')?.textContent||'Metric')+' — Source Records');return;}const ranking=e.target.closest('#charts .ranking-item');if(ranking&&!ranking.matches('.pr41-evidence-clickable, .ip-combined-org-row')){const name=ranking.querySelector('.ranking-name')?.textContent;if(!name)return;const h=ranking.closest('.panel')?.querySelector('h3')?.textContent||'';let kind=h.includes('Department')?'Department':h.includes('Faculty')?'Faculty':h.includes('Institution')?'Organization / Partner':null;open(name+' — Source Records',kind?(r,ctx)=>String(val(r,kind,ctx)).trim()===name.trim():null);}});
// Native Chart.js charts: click a bar/segment to inspect the corresponding source records.
setInterval(()=>{document.querySelectorAll('#charts canvas').forEach(canvas=>{if(canvas.dataset.uolPopupBound)return;const chart=window.Chart?.getChart?.(canvas);if(!chart)return;canvas.dataset.uolPopupBound='1';canvas.classList.add('uol-popup-clickable');canvas.addEventListener('click',e=>{const hits=chart.getElementsAtEventForMode(e,'nearest',{intersect:true},true);if(!hits.length)return;const name=String(chart.data.labels?.[hits[0].index]??'');const title=canvas.closest('.panel')?.querySelector('h3')?.textContent||'';const kind=/Country/i.test(title)?'Country':/Type|Activity/i.test(title)?'Activity':/Faculty/i.test(title)?'Faculty':/Department/i.test(title)?'Department':/Status/i.test(title)?'Status':null;open(name+' — Records',kind?(r,ctx)=>{if(kind==='Country')return r.Country===name;if(kind==='Status')return r.Status===name;return String(val(r,kind,ctx))===name}:null)})})},800);
})();
