
(()=>{
const dashboard=document.getElementById('filterGraphDashboard');
if(!dashboard)return;
let instances=[];let modalChart=null;let requestId=0;let lastSpecs=[];
const palette=['#0b8f73','#2f6cc7','#875acb','#d99737','#2090a7','#cb657f','#5a7fd0','#2db294','#7b88d1','#f1b55a','#6a91a7','#9d75d9'];
const graphTones={faculty:['#2d74da','#9bc1f1'],activity:['#0b8f73','#8ed6c5'],department:['#7a5bd0','#cfbfef'],semester:['#d68c29','#efd59a']};
const clean=v=>String(v??'').trim();
const short=(v,n=24)=>v.length>n?v.slice(0,n-1)+'…':v;
const pretty=n=>Number(n||0).toLocaleString();
function colorPair(spec){return graphTones[spec.key]||['#0b8f73','#9bd9c9'];}
function colorsFor(spec){
  if(spec.type==='doughnut') return spec.items.map((_,i)=>palette[i%palette.length]);
  return (ctx)=>{
    const {chart}=ctx; const area=chart.chartArea; const [deep,light]=colorPair(spec);
    if(!area) return deep;
    const g=chart.ctx.createLinearGradient(area.left,0,area.right,0);
    if(chart.options.indexAxis==='y'){g.addColorStop(0,light);g.addColorStop(.65,deep);g.addColorStop(1,deep);} 
    else {const gv=chart.ctx.createLinearGradient(0,area.bottom,0,area.top); gv.addColorStop(0,light); gv.addColorStop(1,deep); return gv;}
    return g;
  };
}
const chartBackdrop={id:'uolChartBackdrop',beforeDraw(chart){const {ctx,chartArea}=chart; if(!chartArea) return; ctx.save(); const {left,top,width,height}=chartArea; const r=12; ctx.fillStyle='#f8fbfd'; ctx.strokeStyle='#eef3f6'; ctx.lineWidth=1; ctx.beginPath(); ctx.moveTo(left+r,top); ctx.arcTo(left+width,top,left+width,top+height,r); ctx.arcTo(left+width,top+height,left,top+height,r); ctx.arcTo(left,top+height,left,top,r); ctx.arcTo(left,top,left+width,top,r); ctx.closePath(); ctx.fill(); ctx.stroke(); ctx.restore();}};
const professionalValues={id:'uolProfessionalValueLabels',afterDatasetsDraw(chart){if(chart.config.type!=='bar')return;const {ctx,chartArea}=chart;const meta=chart.getDatasetMeta(0);if(!meta||!meta.data)return;ctx.save();ctx.font='700 10px Segoe UI, Arial';ctx.fillStyle='#435c70';ctx.textBaseline='middle';meta.data.forEach((bar,i)=>{const value=Number(chart.data.datasets[0].data[i]||0);if(chart.options.indexAxis==='y'){ctx.textAlign='left';const x=Math.min(bar.x+7,chartArea.right-40);ctx.fillText(pretty(value),x,bar.y)}else{ctx.textAlign='center';ctx.fillText(pretty(value),bar.x,Math.max(chartArea.top+12,bar.y-11))}});ctx.restore()}};
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const total=(list)=>list.reduce((s,p)=>s+(Number(p[1])||0),0);
function dataset(){try{return Array.isArray(filteredRows)?filteredRows:[]}catch(e){return []}}
function aggregate(rows,key){const map=new Map();for(const r of rows){const label=clean(key(r));if(!label||label==='—')continue;const v=Number(rowAmount(r))||0;map.set(label,(map.get(label)||0)+v)}return [...map.entries()].sort((a,b)=>b[1]-a[1])}
function chartOptions(type,horizontal,items){
 const doughnut=type==='doughnut';
 return {
   responsive:true,maintainAspectRatio:false,animation:{duration:350,easing:'easeOutQuart'},
   indexAxis:horizontal?'y':'x',
   cutout:doughnut?'67%':undefined,
   interaction:{mode:'nearest',intersect:true},
   plugins:{
     legend:{display:doughnut,position:'bottom',align:'center',labels:{usePointStyle:true,pointStyle:'circle',boxWidth:8,boxHeight:8,padding:12,color:'#556d80',font:{family:'Segoe UI',size:10,weight:'600'},generateLabels:chart=>chart.data.labels.map((label,i)=>({text:`${label} · ${pretty(chart.data.datasets[0].data[i])}`,fillStyle:Array.isArray(chart.data.datasets[0].backgroundColor)?chart.data.datasets[0].backgroundColor[i]:palette[i%palette.length],strokeStyle:Array.isArray(chart.data.datasets[0].backgroundColor)?chart.data.datasets[0].backgroundColor[i]:palette[i%palette.length],index:i,hidden:false}))}},
     tooltip:{backgroundColor:'#173b56',titleColor:'#ffffff',bodyColor:'#e9f5ff',padding:12,cornerRadius:10,displayColors:true,caretPadding:9,titleFont:{weight:'700'},bodyFont:{weight:'600'},callbacks:{title:ctx=>items[ctx[0]?.dataIndex]?.[0]||'',label:ctx=>`Reported activities: ${pretty(ctx.raw)}`}}
   },
   layout:{padding:{top:22,right:horizontal?60:18,left:8,bottom:10}},
   scales:doughnut?{}:{
     x:{beginAtZero:horizontal,border:{display:false},grid:{display:!horizontal,color:'#eaf0f4',drawTicks:false},ticks:{font:{family:'Segoe UI',size:10,weight:'600'},color:'#73879a',padding:9,maxRotation:0,minRotation:0,precision:0}},
     y:{beginAtZero:!horizontal,border:{display:false},grid:{display:horizontal?false:true,color:'#eaf0f4',drawTicks:false},ticks:{font:{family:'Segoe UI',size:10,weight:'600'},color:'#60778b',padding:9,autoSkip:!horizontal,precision:0}}
   }
 };
}

function rowsForSpec(spec,label){
  const rows=dataset();
  if(!label)return rows;
  return rows.filter(r=>{
    if(spec.key==='semester') return clean(semesterForActivity(r))===clean(label);
    if(spec.key==='activity') return clean(r.type)===clean(label);
    return clean(r[spec.key])===clean(label);
  });
}
function summarizeCard(article,items){
  const chip=article.querySelector('.filter-graph-chip');
  const note=article.querySelector('.filter-graph-note strong');
  if(chip) chip.textContent=`${items.length} groups`;
  if(note) note.textContent=`${total(items).toLocaleString()} reported activities`;
}
function ensureCardChrome(article,title){
  if(article.dataset.enhanced==='1') return;
  article.dataset.enhanced='1';
  const h3=article.querySelector('h3');
  const head=document.createElement('div');
  head.className='filter-graph-cardhead';
  head.innerHTML=`<div class="filter-graph-cardcopy"><h3>${esc(title||h3?.textContent||'Graph')}</h3><p>Click a bar or segment to view filtered supporting records.</p></div><div class="filter-graph-cardmeta"><span class="filter-graph-chip">0 groups</span><button type="button" class="filter-graph-open">Open popup</button></div>`;
  article.prepend(head);
  h3?.remove();
  const note=document.createElement('div');
  note.className='filter-graph-note';
  note.innerHTML='<span>Interactive chart view</span><strong>0 reported activities</strong>';
  article.appendChild(note);
}
function renderModal(spec,selectedLabel){
  const modal=document.getElementById('filterGraphModal'); if(!modal) return;
  const rows=rowsForSpec(spec,selectedLabel);
  const items=spec.items;
  document.getElementById('filterGraphModalTitle').textContent=selectedLabel?`${spec.title} — ${selectedLabel}`:spec.title;
  document.getElementById('filterGraphModalMeta').textContent=selectedLabel?'Showing records behind the selected chart item.':'Showing the full graph for the current filter selection.';
  document.getElementById('filterGraphModalBadge').textContent=`${rows.length.toLocaleString()} rows`;
  document.getElementById('filterGraphModalValue').textContent=`${rows.reduce((sum,r)=>sum+(Number(rowAmount(r))||0),0).toLocaleString()} reported activities`;
  document.getElementById('filterGraphModalRecordsCount').textContent=`${rows.length.toLocaleString()} records`;
  const tbody=document.getElementById('filterGraphModalTbody');
  tbody.innerHTML=rows.slice(0,150).map(r=>`<tr><td>${esc(r.ref||r.refId||r.id||'—')}</td><td>${esc(r.faculty||'—')}</td><td>${esc(r.department||'—')}</td><td>${esc(r.type||'—')}</td><td>${esc(r.organization||'—')}</td><td>${(Number(rowAmount(r))||0).toLocaleString()}</td></tr>`).join('') || '<tr><td colspan="6">No matching records found.</td></tr>';
  const canvas=document.getElementById('filterGraphModalCanvas'); const popupCanvasHolder=canvas?.parentElement; if(popupCanvasHolder)popupCanvasHolder.style.height=spec.key==='activity'?Math.max(420,110+spec.items.length*37)+'px':'420px';
  if(modalChart){try{modalChart.destroy()}catch(e){}}
  modal.hidden=false; document.body.style.overflow='hidden';
  modalChart=new Chart(canvas,{type:spec.type,data:{labels:items.map(p=>short(p[0],32)),datasets:[{label:'Reported activities',data:items.map(p=>p[1]),backgroundColor:colorsFor(spec),borderColor:spec.type==='doughnut'?'#ffffff':'transparent',borderWidth:spec.type==='doughnut'?2:0,borderRadius:spec.type==='bar'?8:0,borderSkipped:false,maxBarThickness:33,hoverBackgroundColor:Array.isArray(graphTones[spec.key])?graphTones[spec.key][0]:graphTones[spec.key]}]},options:chartOptions(spec.type,spec.horizontal,items),plugins:[chartBackdrop, professionalValues]});
}
function closeModal(){const modal=document.getElementById('filterGraphModal'); if(!modal) return; modal.hidden=true; document.body.style.overflow=''; if(modalChart){try{modalChart.destroy()}catch(e){} modalChart=null;}}
function draw(){
  if(dashboard.hidden||getComputedStyle(dashboard).display==='none')return;
  if(typeof Chart==='undefined')return;
  const rows=dataset();
  const faculty=aggregate(rows,r=>r.faculty).slice(0,10),activity=aggregate(rows,r=>r.type),dept=aggregate(rows,r=>r.department).slice(0,12),semester=aggregate(rows,r=>semesterForActivity(r));
  document.getElementById('filterGraphCount').textContent=`${rows.length.toLocaleString()} source records · ${rows.reduce((sum,r)=>sum+(Number(rowAmount(r))||0),0).toLocaleString()} reported activities`;
  instances.forEach(c=>{try{c.destroy()}catch(e){}});instances=[];
  const specs=[
    {id:'filterFacultyGraph',title:'Faculty-wise Activity Volume',type:'bar',items:faculty,horizontal:true,key:'faculty'},
    {id:'filterActivityGraph',title:'All Activity Types',type:'bar',items:activity,horizontal:true,key:'activity'},
    {id:'filterDepartmentGraph',title:'Department-wise Activity Volume',type:'bar',items:dept,horizontal:true,key:'department'},
    {id:'filterSemesterGraph',title:'Semester Comparison',type:'bar',items:semester,horizontal:false,key:'semester'}
  ];
  lastSpecs=specs;
  for(const spec of specs){
    const canvas=document.getElementById(spec.id); if(!canvas) continue;
    const article=canvas.closest('article'); ensureCardChrome(article,spec.title); summarizeCard(article,spec.items); if(spec.key==='activity'){const holder=article.querySelector('.filter-graph-canvas'); if(holder)holder.style.height=Math.max(300,120+spec.items.length*37)+'px';} 
    const labels=spec.items.map(p=>short(p[0],spec.horizontal?32:22)); const values=spec.items.map(p=>p[1]);
    const chart=new Chart(canvas,{type:spec.type,data:{labels,datasets:[{label:'Reported activities',data:values,backgroundColor:colorsFor(spec),borderColor:spec.type==='doughnut'?'#ffffff':'transparent',borderWidth:spec.type==='doughnut'?2:0,borderRadius:spec.type==='bar'?8:0,borderSkipped:false,maxBarThickness:29,hoverBackgroundColor:Array.isArray(graphTones[spec.key])?graphTones[spec.key][0]:graphTones[spec.key]}]},options:chartOptions(spec.type,spec.horizontal,spec.items),plugins:[chartBackdrop, professionalValues]});
    canvas.onclick=(evt)=>{const pts=chart.getElementsAtEventForMode(evt,'nearest',{intersect:true},true); if(pts.length){const idx=pts[0].index; renderModal(spec,spec.items[idx]?.[0]||null);} else {renderModal(spec,null);} };
    article.querySelector('.filter-graph-open').onclick=()=>renderModal(spec,null);
    instances.push(chart);
  }
}
function refresh(){cancelAnimationFrame(requestId);requestId=requestAnimationFrame(()=>requestAnimationFrame(draw))}
const observer=new MutationObserver(refresh); const metric=document.getElementById('kActivities'); if(metric)observer.observe(metric,{childList:true,characterData:true,subtree:true});
for(const el of document.querySelectorAll('#top-filters select,#top-filters input')){el.addEventListener('change',refresh);el.addEventListener('input',refresh)}
const nav=document.querySelector('.executive-on-demand'); if(nav)nav.addEventListener('click',e=>{if(e.target.closest('[data-exec-view="filters"]'))setTimeout(refresh,30)});
document.getElementById('filterGraphExport').addEventListener('click',()=>{
  refresh(); const canvases=[...dashboard.querySelectorAll('canvas')]; if(!instances.length)return;
  const w=1600,h=1100; const output=document.createElement('canvas'); output.width=w; output.height=h; const ctx=output.getContext('2d'); ctx.fillStyle='#f6fbfd'; ctx.fillRect(0,0,w,h); ctx.fillStyle='#17364e'; ctx.font='bold 27px Arial'; ctx.fillText('UOL — Filtered Analytics',35,47); ctx.font='16px Arial'; ctx.fillText(document.getElementById('filterGraphCount').textContent,35,75);
  canvases.forEach((canvas,i)=>{const x=35+(i%2)*785,y=105+Math.floor(i/2)*490; ctx.fillStyle='#ffffff'; ctx.strokeStyle='#dfeaf1'; ctx.lineWidth=1; ctx.fillRect(x,y,745,455); ctx.strokeRect(x,y,745,455); ctx.fillStyle='#17364e'; ctx.font='bold 18px Arial'; const title=canvas.closest('article').querySelector('.filter-graph-cardcopy h3').textContent; ctx.fillText(title,x+14,y+28); try{ctx.drawImage(canvas,x+14,y+43,716,390)}catch(e){}});
  const a=document.createElement('a'); a.download='UOL_Filtered_Analytics.png'; a.href=output.toDataURL('image/png'); a.click();
});
const modal=document.getElementById('filterGraphModal'); if(modal){modal.addEventListener('click',e=>{if(e.target.matches('[data-close-filter-graph-modal], .filter-graph-modal')) closeModal();}); document.getElementById('filterGraphModalClose')?.addEventListener('click',closeModal); document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!modal.hidden) closeModal();});}
refresh();
})();
