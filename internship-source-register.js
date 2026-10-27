(()=>{'use strict';
const host=document.getElementById('internshipSourceRegister');if(!host)return;
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const clean=s=>String(s??'').replace(/\s+/g,' ').trim();
const file='data/internship-evidence-user-20261027.json';
fetch(file,{cache:'no-store'}).then(r=>{if(!r.ok)throw Error('HTTP '+r.status);return r.json()}).then(data=>{
 if(!Array.isArray(data.records))throw Error('Missing records array in JSON.');
 const rows=data.records.map((r,i)=>({
  index:i,ref:clean(r.ref),submittedBy:clean(r.submittedBy),status:clean(r.status),submissionDate:clean(r.submissionDate),month:clean(r.month),campus:clean(r.campus),faculty:clean(r.faculty),department:clean(r.department),person:clean(r.person),designation:clean(r.designation),type:clean(r.type),quantity:Number(r.quantity)||0,organization:clean(r.organization),planned:clean(r.planned),scheduled:clean(r.scheduled),remarks:clean(r.remarks)
 })).filter(r=>r.organization && !/^\.*$/.test(r.organization));
 const validOrg=r=>r.organization && !/^\.*$/.test(r.organization);
 const unique=key=>[...new Set(rows.map(r=>r[key]).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
 const makeFilter=(id,title,values)=>`<label class="ip-source-filter">${title}<select id="${id}"><option value="">All ${title}</option>${values.map(s=>`<option value="${esc(s)}">${esc(s)}</option>`).join('')}</select></label>`;
 host.innerHTML=`<div class="ip-source-head"><div><span class="ip-source-kicker">LIVE JSON EVIDENCE SOURCE</span><h2>Internship / Placement — Organization Evidence</h2><p>Connected file: <strong>internship-evidence-user-20261027.json</strong>. Filter by organization or activity type, or search by name and remarks. Click any organization for its documented evidence. The confirmed submitted-list quantity is <strong>127</strong>; this JSON contains <strong>117 named organization records</strong>, leaving 10 entries without matching organization-level details. Faculty, department and submission references are unavailable in this JSON.</p></div><span class="ip-source-badge" id="ipSourceBadge">${rows.length} records</span></div>
 <div class="ip-source-tools ip-source-filters">${makeFilter('ipSrcType','Activity Types',unique('type'))}${makeFilter('ipSrcOrg','Organizations',unique('organization'))}<label class="ip-source-filter">Search<input type="search" id="ipSourceSearch" placeholder="Find company, Ref ID, remarks…" aria-label="Search internship worksheet"></label><button type="button" id="ipSrcReset" class="ip-source-reset">Reset</button></div><div class="ip-source-stats"><strong id="ipSourceVisible"></strong><span id="ipSourceQty"></span></div><div class="ip-source-list" id="ipSourceList"></div><p class="ip-source-note">This register reads the JSON file live on both dashboard pages. Evidence quantities are shown separately and are not added to the institution-wide student total to avoid double-counting. Submission references, faculties, and departments are not included in this JSON.</p>`;
 const modal=document.createElement('div');modal.className='ip-source-modal';modal.hidden=true;modal.innerHTML='<div class="ip-source-dialog" role="dialog" aria-modal="true" aria-labelledby="ipSourceModalTitle"><div class="ip-source-dialog-head"><div><small>INTERNSHIP / PLACEMENT — JSON EVIDENCE</small><h2 id="ipSourceModalTitle"></h2></div><button type="button" class="ip-source-close" aria-label="Close">×</button></div><div class="ip-source-dialog-body" id="ipSourceModalBody"></div></div>';document.body.append(modal);
 const close=()=>{modal.hidden=true;document.body.style.overflow='';};modal.querySelector('.ip-source-close').addEventListener('click',close);modal.addEventListener('click',e=>{if(e.target===modal)close();});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!modal.hidden)close()});
 const $=id=>host.querySelector('#'+id);const controlIds=['ipSrcType','ipSrcOrg'];
 const parentControls=['faculty','department','organization'];
 function render(){
  const term=$('ipSourceSearch').value.toLowerCase().trim();
  const subset=rows.filter(r=>controlIds.every((id)=>{const v=$(id).value;return !v||r[({'ipSrcType':'type','ipSrcOrg':'organization'})[id]]===v}) && (!term||[r.type,r.organization,r.remarks].join(' ').toLowerCase().includes(term)));
  $('ipSourceVisible').textContent=`${subset.length} / ${rows.length} matching source rows`;
  $('ipSourceQty').textContent=`${subset.reduce((v,r)=>v+r.quantity,0).toLocaleString()} JSON evidence quantity (not added again to KPI)`;
  $('ipSourceList').innerHTML=subset.map(r=>`<button type="button" class="ip-source-item" data-source-index="${r.index}"><span class="ip-source-name"><strong>${esc(validOrg(r)?r.organization:'Organization not provided')}</strong><small>${esc(r.type)} · JSON evidence record</small></span><span class="ip-source-count">${r.quantity} ${r.quantity===1?'entry':'entries'}</span></button>`).join('')||'<p>No matching records in the separate worksheet.</p>';
  $('ipSourceList').querySelectorAll('[data-source-index]').forEach(b=>b.addEventListener('click',()=>{
   const r=rows[Number(b.dataset.sourceIndex)];modal.querySelector('#ipSourceModalTitle').textContent=validOrg(r)?r.organization:r.ref+' — Organization unspecified';
   const details=[['Recorded quantity',r.quantity],['Activity type',r.type],['Planned',r.planned],['Scheduled',r.scheduled],['Remarks / result',r.remarks]];
   modal.querySelector('#ipSourceModalBody').innerHTML=details.map(([k,v])=>`<div class="ip-source-detail"><span>${esc(k)}</span><strong>${esc(v||'Not provided')}</strong></div>`).join('');modal.hidden=false;document.body.style.overflow='hidden';
  }));
 }
 for(const id of controlIds)$(id).addEventListener('change',render);
 $('ipSourceSearch').addEventListener('input',render);
 $('ipSrcReset').addEventListener('click',()=>{for(const id of controlIds)$(id).value='';$('ipSourceSearch').value='';render();});
 render();
}).catch(e=>{host.textContent='Unable to load the internship JSON evidence: '+e.message+' — Run START_CONNECTED_DASHBOARD.bat and check data/internship-evidence-user-20261027.json.';});
})();
