(()=>{'use strict';
// PRPSTFPT41 organizations are exposed ONLY in the regular organization dropdowns.
// No independent audit table or toolbar is rendered.
const DATA='data/PRPSTFPT41_deduplication_audit.json';
const faculty='Faculty of Information Technology',department='Department of Software Engineering';
const section2=document.body.classList.contains('ip-modern');
const id=section2?'organization':'organizationFilter';
const get=(key)=>document.getElementById(key)?.value||'All';
const match=(value,target)=>value==='All'||!value||value===target;
fetch(DATA,{cache:'no-store'}).then(r=>{if(!r.ok)throw Error(r.status);return r.json()}).then(data=>{
 const el=document.getElementById(id);if(!el)return;
 let busy=false,refreshTimer;
 const enhance=()=>{
  if(busy)return;busy=true;
  try{
   const f=get(section2?'faculty':'facultyFilter'),d=get(section2?'department':'departmentFilter'),a=section2?'All':get('activityFilter');
   const selected=el.value;
   const allowed=match(f,faculty)&&match(d,department)&&(a==='All'||/internship|placement|^ip$/i.test(a));
   for(const opt of Array.from(el.querySelectorAll('option[data-pr41-audit]')))opt.remove();
   for(const group of Array.from(el.querySelectorAll('optgroup[data-pr41-group]')))group.remove();
   if(allowed){
     const existing=new Set(Array.from(el.options).map(o=>o.value.toLowerCase()));
     const group=document.createElement('optgroup');group.label='Software Engineering · PRPSTFPT41';group.dataset.pr41Group='1';
     for(const item of data.organizations||[]){if(existing.has(item.name.toLowerCase()))continue;
       const option=new Option(`${item.name} (${item.evidenceRecords} evidence records)`,item.name);
       option.dataset.pr41Audit='1';group.appendChild(option);
     }
     if(group.children.length)el.appendChild(group);
     if(selected!=='All'&&Array.from(el.options).some(o=>o.value===selected))el.value=selected;
   }
  }finally{busy=false;}
 };
 const observe=new MutationObserver(()=>{if(!busy){clearTimeout(refreshTimer);refreshTimer=setTimeout(enhance,100);}});
 observe.observe(el,{childList:true,subtree:true});
 const keys=section2?['faculty','department','semester','reset']:['facultyFilter','departmentFilter','activityFilter','monthFilter'];
 keys.forEach(key=>document.getElementById(key)?.addEventListener('change',()=>setTimeout(enhance,80)));
 enhance();
}).catch(err=>console.warn('Organization audit options unavailable',err));
})();
