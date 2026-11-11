/* Hide the Semester field in detail overlays on Sections 1–3.
   This is presentation-only.  Raw workbook data is untouched. */
(() => {
  'use strict';
  const exact = text => /^semester(?:\s*:\s*)?$/i.test(String(text || '').trim());
  const insidePopup = el => el.closest('dialog,[role="dialog"],.modal,.popup,.uol-major-overlay,.pr41-company-overlay,.overlay,[class*="-modal"],[class*="-popup"]');
  function cleanPopup(popup) {
    if (!popup || popup.nodeType !== 1) return;
    // Two-column field cards and definition lists.
    popup.querySelectorAll('dt,th,label,.detail-label,.field-label,.modal-label,.record-label,small,strong,b,span').forEach(el => {
      if (!exact(el.textContent) || el.children.length || !insidePopup(el)) return;
      const card = el.closest('.detail-item,.detail-field,.record-field,.modal-field,.field-row,.field-card,.detail-card,.record-card,dl > div');
      if (card && card !== popup) { card.remove(); return; }
      if (el.tagName === 'DT') { const next=el.nextElementSibling; if(next?.tagName==='DD') next.remove(); el.remove(); return; }
      if (el.tagName === 'TH') {
        const table=el.closest('table'); const index=el.cellIndex;
        if(table && index >= 0) {
          table.querySelectorAll('tr').forEach(row => row.cells[index]?.remove());
        }
        return;
      }
      const parent = el.parentElement;
      if(parent && parent !== popup && parent.children.length === 2 && parent.querySelector('small,span,strong,b,label')) parent.remove();
    });
  }
  function clean() {
    document.querySelectorAll('dialog,[role="dialog"],.modal,.popup,.uol-major-overlay,.pr41-company-overlay,.overlay,[class*="-modal"],[class*="-popup"]').forEach(cleanPopup);
  }
  let working=false;
  const observer=new MutationObserver(() => {
    if(working)return;
    working=true;
    queueMicrotask(() => { try { clean(); } finally { working=false; } });
  });
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded',()=>{clean();observer.observe(document.body,{childList:true,subtree:true});},{once:true});
  } else {clean();observer.observe(document.body,{childList:true,subtree:true});}
})();
