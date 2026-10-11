/* Shared mobile navigation; no changes to charts or reporting calculations. */
(() => {
  const sidebar = document.querySelector('#sidebar, .uol-shared-sidebar');
  const button = document.querySelector('#mobileNavToggle, .uol-mobile-menu-button');
  if (!sidebar || !button) return;
  const overview = !!document.querySelector('#mobileNavToggle');
  const menu = document.getElementById(overview ? 'navMenu' : 'uolMobileNavigation');
  if (!menu) return;
  const setOpen = value => {
    sidebar.classList.toggle(overview ? 'mobile-open' : 'uol-mobile-open', value);
    button.setAttribute('aria-expanded', String(value));
    button.setAttribute('aria-label', value ? 'Close dashboard navigation' : 'Open dashboard navigation');
    const chevron = button.querySelector('.uol-menu-chevron');
    if (chevron) chevron.textContent = value ? '⌃' : '⌄';
  };
  if (overview) {
    /* app.js owns the existing Overview click handler; sync its accessibility state afterwards. */
    button.addEventListener('click', () => queueMicrotask(() => {
      button.setAttribute('aria-expanded', String(sidebar.classList.contains('mobile-open')));
    }));
  } else {
    button.addEventListener('click', () => setOpen(!sidebar.classList.contains('uol-mobile-open')));
  }
  menu.addEventListener('click', event => {
    if (event.target.closest('a[href]') && matchMedia('(max-width: 1050px)').matches) setOpen(false);
  });
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape') setOpen(false);
  });
  document.addEventListener('click', event => {
    if (matchMedia('(max-width: 1050px)').matches && !sidebar.contains(event.target)) setOpen(false);
  });
  const query = matchMedia('(min-width: 1051px)');
  query.addEventListener?.('change', event => { if (event.matches) setOpen(false); });
})();
