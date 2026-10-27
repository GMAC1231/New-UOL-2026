(()=>{
 const sidebar=document.querySelector('#sidebar, .uol-shared-sidebar');
 if(!sidebar)return;
 const old=sidebar.querySelector('.mobile-nav-toggle, .uol-mobile-menu-button');
 if(old){old.disabled=true;old.setAttribute('aria-hidden','true')}
 const menu=sidebar.querySelector('#navMenu, #uolMobileNavigation');
 if(!menu)return;
 const button=document.createElement('button');button.type='button';
 button.className='uol-global-hamburger';button.setAttribute('aria-controls',menu.id);
 button.innerHTML='<span class="uol-three-bars" aria-hidden="true"><i></i><i></i><i></i></span><span class="uol-hamburger-label">Dashboard Menu</span>';
 sidebar.insertBefore(button,sidebar.firstChild);
 let isOpen=window.innerWidth>1050;
 function apply(open){isOpen=open;document.body.classList.toggle('uol-sidebar-open',open);button.setAttribute('aria-expanded',String(open));button.setAttribute('aria-label',open?'Collapse dashboard menu':'Expand dashboard menu');button.title=open?'Collapse menu':'Expand menu';window.dispatchEvent(new Event('resize'))}
 button.addEventListener('click',()=>apply(!isOpen));
 menu.addEventListener('click',e=>{if(e.target.closest('a[href]')&&window.innerWidth<=1050)apply(false)});
 document.addEventListener('keydown',e=>{if(e.key==='Escape'&&window.innerWidth<=1050)apply(false)});
 window.addEventListener('resize',()=>{if(window.innerWidth<=1050 && !isOpen) return;});
 apply(isOpen);
})();
