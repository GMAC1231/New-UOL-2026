(() => {
const links=[['index.html','Executive Overview'],['basic-reporting.html','Section 1 · Basic Reporting'],['internship-placement.html','Section 2 · Internship & Placement'],['mou-collaborations.html','Section 3 · MoU & Collaborations']];
const active=location.pathname.split('/').pop()||'index.html';
if(active==='index.html')return;
const target=document.querySelector('.report-inner, .basic-content');
if(!target)return;
const nav=document.createElement('div');nav.className='uol-section-jump';
const lab=document.createElement('label');lab.textContent='Navigate dashboard';lab.htmlFor='uol-section-select';
const select=document.createElement('select');select.id='uol-section-select';select.setAttribute('aria-label','Navigate reporting sections');
for(const [url,title] of links){const opt=document.createElement('option');opt.value=url;opt.textContent=title;opt.selected=active===url;select.appendChild(opt)}
select.addEventListener('change',()=>{location.href=select.value});nav.append(lab,select);target.prepend(nav);
})();
