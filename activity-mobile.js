(() => {
 const sidebar=document.querySelector('.side');
 const button=document.querySelector('.activity-mobile-menu-button');
 const menu=document.querySelector('.activity-mobile-navigation');
 if(!sidebar||!button||!menu)return;
 const setOpen=open=>{sidebar.classList.toggle('activity-mobile-open',open);button.setAttribute('aria-expanded',String(open));button.setAttribute('aria-label',open?'Close activity dashboard navigation':'Open activity dashboard navigation');};
 button.addEventListener('click',()=>setOpen(!sidebar.classList.contains('activity-mobile-open')));
 menu.addEventListener('click',e=>{if(e.target.closest('a[href]'))setOpen(false);});
 document.addEventListener('keydown',e=>{if(e.key==='Escape')setOpen(false);});
 document.addEventListener('click',e=>{if(!sidebar.contains(e.target))setOpen(false);});
 matchMedia('(min-width:1051px)').addEventListener?.('change',e=>{if(e.matches)setOpen(false);});
})();
