(()=>{
'use strict';
const pages=[['index.html','Executive Overview'],['basic-reporting.html','Section 1 · Basic Reporting'],['internship-placement.html','Section 2 · Internship / Placement'],['mou-collaborations.html','Section 3 · MoU & Collaborations']];
const active=location.pathname.split('/').pop()||'index.html';
const main=document.querySelector('main');if(!main)return;
// Replace older section-only dropdown, which otherwise introduces duplicate selectors.
document.querySelector('.uol-section-jump')?.remove();
const bar=document.createElement('nav');bar.className='uol-unified-nav';bar.setAttribute('aria-label','Dashboard sections');
const links=document.createElement('div');links.className='uol-unified-links';
for(const [file,label] of pages){let a=document.createElement('a');a.href=file;a.textContent=label;if(file===active){a.setAttribute('aria-current','page');a.classList.add('active')}links.append(a)}
const select=document.createElement('select');select.setAttribute('aria-label','Choose dashboard section');select.className='uol-unified-select';
for(const [file,label] of pages){const o=document.createElement('option');o.value=file;o.textContent=label;o.selected=file===active;select.append(o)}
select.onchange=()=>{location.href=select.value};bar.append(links,select);
const host=main.querySelector('.report-inner')||main;
const before=host.querySelector('.page-intro,.pbi-toolbar,.tentative-section-nav,.filter-strip,.basic-filterbar')||host.firstElementChild;
if(before)host.insertBefore(bar,before);else host.prepend(bar);
// Table rows remain accessible on small screens by horizontal swiping.
for(const table of document.querySelectorAll('table')){if(!table.parentElement?.classList.contains('uol-table-responsive')&&!/table-scroll|tablewrap|uol-popup-table-wrap/.test(table.parentElement?.className||'')){const wrap=document.createElement('div');wrap.className='uol-table-responsive';table.parentNode.insertBefore(wrap,table);wrap.append(table)}}
})();
