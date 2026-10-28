(()=>{
  const charts=document.getElementById('analytics-section');
  const quality=document.getElementById('data-quality-scope');
  const filters=document.getElementById('top-filters');
  const hint=document.getElementById('execViewHint');
  const filterGraphs=document.getElementById('filterGraphDashboard');
  const nav=document.querySelector('.executive-on-demand');
  if(!charts || !nav) return;
  const panelByView={faculty:'.panel-faculty',activity:'.panel-activity',monthly:'.panel-monthly',department:'.panel-department',internship:'.panel-internship',status:'.panel-status'};
  charts.classList.add('exec-view-compact');
  function activate(view,scroll){
    nav.querySelectorAll('[data-exec-view]').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.execView===view)));
    const panel=panelByView[view] && charts.querySelector(panelByView[view]);
    charts.classList.toggle('is-open',!!panel);
    charts.querySelectorAll('.panel').forEach(p=>p.classList.toggle('exec-panel-active',p===panel));
    if(quality)quality.classList.toggle('exec-hidden',view!=='quality');
    if(filters)filters.classList.remove('exec-hidden');
    if(filterGraphs)filterGraphs.classList.remove('exec-visible');
    if(hint)hint.textContent=view==='summary'?'Choose a button above to open one detailed section at a time.':view==='quality'?'Source data audit and quality checks.':'Showing one chart section. Choose Overview to return to the summary.';
    requestAnimationFrame(()=>{window.dispatchEvent(new Event('resize'));if(window.Chart){for(const instance of Object.values(Chart.instances||{})){try{if(instance.canvas?.getClientRects().length)instance.resize()}catch(_){}}}});
    if(scroll && view!=='summary') (view==='quality'?quality:charts)?.scrollIntoView({block:'start',behavior:'smooth'});
  }
  nav.addEventListener('click',e=>{const button=e.target.closest('[data-exec-view]');if(button)activate(button.dataset.execView,true)});
  for(const a of document.querySelectorAll('a[href="#analytics-section"]')){
    a.addEventListener('click',e=>{e.preventDefault();activate('faculty',true)})
  }
  activate(location.hash==='#analytics-section'?'faculty':'summary',false);
})();
