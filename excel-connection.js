/* UOL Excel Data Connection: browser-side overrides with bundled workbook fallbacks. */
(() => {
  'use strict';
  const sources = {
    submissions: {title:'Basic Reporting — 64 Submissions', file:'64_Submissions_Source.xlsx', sheets:['All Submissions']},
    activities: {title:'Executive Overview & Internship / Placement', file:'UOL_Separated_All_Activity_Data_Cleaned.xlsx', sheets:['All Separated Activities','Internship - Placement']},
    mou: {title:'MoU Collaborations', file:'UOL_MoU_Country_Previous.xlsx', sheets:['MoU Country (No Nulls)']}
  };
  const database = 'UOLExcelWorkbookOverridesV1';
  const store = 'workbooks';
  let dbPromise;
  function openDB() {
    if (!('indexedDB' in window)) return Promise.reject(new Error('Browser storage is unavailable'));
    if (!dbPromise) dbPromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(database, 1);
      request.onupgradeneeded = () => request.result.createObjectStore(store);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    return dbPromise;
  }
  async function stored(key) {
    try {
      const db = await openDB();
      return await new Promise((resolve, reject) => {
        const req = db.transaction(store, 'readonly').objectStore(store).get(key);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      });
    } catch (err) {console.warn('Excel override unavailable:', err); return null;}
  }
  async function save(key, payload) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const req = db.transaction(store, 'readwrite').objectStore(store).put(payload, key);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  }
  async function forget(key) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const req = db.transaction(store, 'readwrite').objectStore(store).delete(key);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  }
  function read(buffer) {return XLSX.read(new Uint8Array(buffer), {type:'array',cellDates:false});}
  function validate(wb,key) {
    const missing = sources[key].sheets.filter(s => !wb.SheetNames.includes(s));
    if (missing.length) throw new Error('Required Excel sheet(s) missing: ' + missing.join(', ') + '. Please choose the correct source workbook.');
    for (const sheet of sources[key].sheets) {
      if (!XLSX.utils.sheet_to_json(wb.Sheets[sheet],{defval:''}).length) throw new Error('Sheet is empty: ' + sheet);
    }
  }
  async function load(key, defaultPath) {
    const override = await stored(key);
    if (override && override.buffer) {
      try {
        const workbook = read(override.buffer);
        validate(workbook,key);
        return {workbook, path: 'Connected Excel: ' + override.name, override:true};
      } catch(e) {console.error('Invalid saved workbook',e); throw e;}
    }
    const paths = Array.isArray(defaultPath) ? defaultPath : [defaultPath];
    let lastError;
    for (const path of paths) {
      try {
        const response = await fetch(path,{cache:'no-store'});
        if(!response.ok) throw new Error(`HTTP ${response.status}: ${path}`);
        const workbook = read(await response.arrayBuffer());
        validate(workbook,key);
        return {workbook, path, override:false};
      } catch(e) {lastError=e; console.warn('Default workbook could not load:',path,e);}
    }
    throw lastError || new Error('Excel file not available. Host the folder on a website and keep the data directory.');
  }
  function el(tag,props={},children=[]) {
    const node=document.createElement(tag);
    Object.entries(props).forEach(([k,v]) => {if(k==='text')node.textContent=v;else if(k==='class')node.className=v;else node.setAttribute(k,v);});
    for (const child of children) node.append(child);
    return node;
  }
  function initUI() {
    const styles = el('style',{text:`
      .excel-connect-open{display:inline-flex;gap:8px;align-items:center;justify-content:center;border:1px solid #c5dbd5;background:#fff;color:#146a59;border-radius:10px;padding:9px 14px;font-size:12px;font-weight:800;cursor:pointer;box-shadow:0 3px 9px #0d3b4a0d;margin:7px 0}
      .excel-connect-open:hover{background:#eef9f5}.excel-connect-backdrop{position:fixed;inset:0;background:#081e33b9;z-index:100020;display:none;align-items:center;justify-content:center;padding:20px}
      .excel-connect-backdrop.open{display:flex}.excel-connect-dialog{width:min(650px,100%);max-height:90vh;overflow-y:auto;background:white;color:#17344c;border-radius:18px;padding:26px;box-shadow:0 24px 80px #051a3473;font:14px/1.55 'Segoe UI',Arial,sans-serif}
      .excel-connect-dialog h2{margin:0;font-size:23px}.excel-connect-dialog p{color:#536d83;font-size:13px;margin:8px 0 15px}.excel-connect-row{padding:14px 0;border-top:1px solid #e4ebf3}.excel-connect-row strong{display:block;margin-bottom:4px}.excel-connect-row small{display:block;color:#6b8093;margin-bottom:8px}
      .excel-connect-row input{width:100%;font-size:12px;max-width:100%;padding:8px;border:1px solid #d3e0ec;border-radius:7px}.excel-connect-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}.excel-connect-actions button,.excel-connect-close{background:#087d67;color:white;border:none;padding:8px 13px;border-radius:8px;cursor:pointer;font-size:12px}.excel-connect-actions button.secondary{color:#24495f;background:#edf3f8}.excel-connect-status{font-size:12px;color:#166c57;min-height:20px;margin-top:6px;overflow-wrap:anywhere}
    `});document.head.append(styles);
    const backdrop=el('div',{class:'excel-connect-backdrop',role:'presentation'});
    const dialog=el('div',{class:'excel-connect-dialog',role:'dialog','aria-modal':'true','aria-label':'Connect Excel data'});
    dialog.append(el('h2',{text:'Connect Excel Data'}),el('p',{text:'Import a compatible workbook for each report. The dashboard reloads with the selected Excel data. Your files remain in this browser; the bundled Excel sheets are used by default.'}));
    Object.entries(sources).forEach(([key,config]) => {
      const row=el('div',{class:'excel-connect-row'});
      const status=el('div',{class:'excel-connect-status',id:'excel-status-'+key});
      const input=el('input',{type:'file',accept:'.xlsx,.xls',id:'excel-file-'+key,'aria-label':'Choose '+config.title+' Excel file'});
      const replace=el('button',{type:'button',text:'Use selected Excel'});
      const reset=el('button',{type:'button',class:'secondary',text:'Restore bundled Excel'});
      row.append(el('strong',{text:config.title}),el('small',{text:'Required sheet(s): '+config.sheets.join(' + ')}),input,el('div',{class:'excel-connect-actions'},[replace,reset]),status);
      replace.addEventListener('click',async () => {
        if(!input.files.length){status.textContent='Choose an Excel file first.';return;}
        replace.disabled=true;status.textContent='Validating Excel workbook…';
        try {
          const file=input.files[0], buffer=await file.arrayBuffer(),wb=read(buffer);
          validate(wb,key);await save(key,{buffer,name:file.name});
          status.textContent='Connected: '+file.name+'. Refreshing dashboard…';
          window.location.reload();
        } catch(e) {status.textContent='Unable to connect: '+e.message;replace.disabled=false;}
      });
      reset.addEventListener('click',async () => {try{await forget(key);window.location.reload();}catch(e){status.textContent='Could not restore default: '+e.message;}});
      stored(key).then(record=>{status.textContent=record ? 'Current source: '+record.name+' (browser override)' : 'Current source: bundled data/'+config.file;});
      dialog.append(row);
    });
    const close=el('button',{type:'button',class:'excel-connect-close',text:'Close'});
    close.onclick=()=>{backdrop.classList.remove('open');};dialog.append(close);backdrop.append(dialog);document.body.append(backdrop);
    backdrop.addEventListener('click',e=>{if(e.target===backdrop)backdrop.classList.remove('open');});
    document.addEventListener('keydown',e=>{if(e.key==='Escape')backdrop.classList.remove('open');});
    const open=el('button',{type:'button',class:'excel-connect-open',text:'📁 Connect Excel'});
    open.onclick=()=>backdrop.classList.add('open');
    const target=document.querySelector('.status-right, .pbi-toolbar-right, .filter-strip, .filters, .topbar') || document.body;
    if (target.classList.contains('filter-strip') || target.classList.contains('filters')) target.prepend(open);
    else target.append(open);
  }
  window.UOLExcel={load,initUI};
  // Automatic bundled Excel loading remains enabled. The manual Excel UI is intentionally disabled.
  // No Connect Excel button is injected on the Executive Overview.
})();
