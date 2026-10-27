(()=>{
const charts=document.getElementById('charts');
if(!charts||!document.body.classList.contains('mou-modern'))return;
let pending=false;const gap=12;
function layout(){
 pending=false;
 const wide=window.matchMedia('(min-width:1001px)').matches;
 if(!wide){charts.style.height='';for(const p of charts.children){p.style.left='';p.style.top=''}return}
 const panels=[...charts.children].filter(e=>e.classList.contains('panel')&&getComputedStyle(e).display!=='none');
 if(!panels.length){charts.style.height='0px';return}
 const colW=(charts.clientWidth-gap)/2;
 let heights=[0,0];
 for(const p of panels){
   const col=heights[0]<=heights[1]?0:1;
   p.style.left=(col*(colW+gap))+'px';
   p.style.top=heights[col]+'px';
   heights[col]+=p.offsetHeight+gap;
 }
 charts.style.height=Math.max(0,...heights)-gap+'px';
}
function schedule(){if(pending)return;pending=true;requestAnimationFrame(layout)}
new MutationObserver(schedule).observe(charts,{childList:true});
const ro=new ResizeObserver(schedule);ro.observe(charts);
let tracked=new Set();
function trackPanels(){
 for(const p of charts.children)if(p.classList.contains('panel')&&!tracked.has(p)){tracked.add(p);ro.observe(p)}
 for(const p of [...tracked])if(!p.isConnected){ro.unobserve(p);tracked.delete(p)}
 schedule();
}
new MutationObserver(trackPanels).observe(charts,{childList:true});
window.addEventListener('resize',schedule,{passive:true});
trackPanels();
})();