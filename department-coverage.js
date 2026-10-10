/* Submission department coverage shared across all four UOL pages.
   This is DISTINCT from approved activity/placement/MoU participation totals. */
(()=>{'use strict';
 const $=id=>document.getElementById(id);
 const normal=x=>String(x??'').trim().toLowerCase();
 const valid=x=>x&&!['null','n/a','nan','undefined','not available','unclassified'].includes(normal(x));
 async function setup(){
  const host=document.querySelector('.report-inner .page-intro') ||
   document.querySelector('.pbi-toolbar') || document.querySelector('.filter-strip');
  if(!host || typeof XLSX==='undefined')return;
  try{
   const response=await fetch('data/64_Submissions_Source.xlsx',{cache:'no-store'});
   if(!response.ok)throw Error('HTTP '+response.status);
   const book=XLSX.read(await response.arrayBuffer(),{type:'array'});
   const sh=book.Sheets['All Submissions'];
   if(!sh)throw Error('All Submissions worksheet missing');
   const all=XLSX.utils.sheet_to_json(sh,{defval:'',raw:false});
   const departments=new Set(all.map(row=>String(row.Department||'').trim()).filter(valid));
   const physics=all.filter(r=>normal(r.Department)==='department of physics');
   const badge=document.createElement('section');badge.className='uol-coverage-banner';
   badge.setAttribute('aria-label','Submission department coverage');
   const main=document.createElement('div');main.className='uol-coverage-main';
   const label=document.createElement('span');label.textContent='INSTITUTIONAL REPORTING COVERAGE';
   const value=document.createElement('strong');value.textContent=departments.size+' Departments';
   main.append(label,value);
   const details=document.createElement('span');details.className='uol-coverage-note';
   details.textContent=all.length+' submissions · Physics: '+physics.length+' submission(s) in source; activity figures are drawn from their respective activity workbooks';
   badge.append(main,details);
   host.insertAdjacentElement('afterend',badge);
  }catch(err){console.warn('Department coverage banner could not load',err);}
 }
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',setup,{once:true});
 else setup();
})();
