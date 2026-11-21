(() => {
  const DataBook = window.DataBook || (window.DataBook = {});
  const navButtons = [...document.querySelectorAll('.nav-item')];
  const views = {
    dashboard: document.getElementById('dashboard-view'),
    'order-summary': document.getElementById('order-summary-view')
  };
  const breadcrumb = document.getElementById('breadcrumb-current');
  const configureMenuWrap = document.getElementById('configure-menu-wrap');
  const configureMenuButton = document.getElementById('configure-menu-button');
  const configureMenu = document.getElementById('configure-menu');
  const openConfigureButton = document.getElementById('open-configure-button');
  const configureModal = document.getElementById('configure-modal');
  const configureDialog = configureModal.querySelector('.configure-dialog');
  const closeConfigureButton = document.getElementById('close-configure-button');
  const configureEmptyState = document.getElementById('configure-empty-state');
  const sheetWorkspace = document.getElementById('sheet-workspace');
  let previousBodyOverflow = '';

  function closeMenu() {
    configureMenu.hidden = true;
    configureMenuButton.setAttribute('aria-expanded', 'false');
  }

  function closeConfigure(restoreFocus = true) {
    if (configureModal.hidden) return;
    configureModal.hidden = true;
    document.body.style.overflow = previousBodyOverflow;
    if (restoreFocus && !configureMenuWrap.hidden) configureMenuButton.focus();
  }

  function openConfigure() {
    closeMenu();
    configureEmptyState.hidden = !sheetWorkspace.hidden;
    previousBodyOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    configureModal.hidden = false;
    closeConfigureButton.focus();
  }

  function showView(name, moveFocus = true) {
    if (!views[name]) return;
    closeConfigure(false);
    closeMenu();
    configureMenuWrap.hidden = name !== 'order-summary';
    Object.entries(views).forEach(([key, view]) => { view.hidden = key !== name; });
    navButtons.forEach(button => {
      const active = button.dataset.view === name;
      button.classList.toggle('active', active);
      if (active) button.setAttribute('aria-current', 'page');
      else button.removeAttribute('aria-current');
    });
    const title = name === 'dashboard' ? 'Dashboard' : 'Order Summary';
    breadcrumb.textContent = title;
    document.title = 'Data Book | ' + title;
    if (location.hash !== '#' + name) history.replaceState(null, '', '#' + name);
    if (moveFocus) views[name].querySelector('h1').focus();
  }

  configureMenuButton.addEventListener('click', () => {
    configureMenu.hidden = !configureMenu.hidden;
    configureMenuButton.setAttribute('aria-expanded', String(!configureMenu.hidden));
  });
  openConfigureButton.addEventListener('click', openConfigure);
  closeConfigureButton.addEventListener('click', () => closeConfigure());
  configureModal.addEventListener('click', event => {
    if (event.target === configureModal) closeConfigure();
  });
  document.addEventListener('click', event => {
    if (!configureMenuWrap.contains(event.target)) closeMenu();
  });
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape') {
      if (!configureModal.hidden) {
        event.preventDefault();
        closeConfigure();
      } else if (!configureMenu.hidden) {
        event.preventDefault();
        closeMenu();
        configureMenuButton.focus();
      }
    }
    if (event.key !== 'Tab' || configureModal.hidden) return;
    const focusable = [...configureDialog.querySelectorAll('button:not([disabled]), input:not([disabled]), select:not([disabled]), summary, [tabindex]:not([tabindex="-1"])')]
      .filter(element => !element.closest('[hidden]') && element.getClientRects().length > 0);
    if (!focusable.length) return;
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  });

  navButtons.forEach(button => button.addEventListener('click', () => showView(button.dataset.view)));
  window.addEventListener('hashchange', () => showView(location.hash.slice(1)));
  DataBook.dashboard.init(showView);
  DataBook.orderSummary.init();
  showView(views[location.hash.slice(1)] ? location.hash.slice(1) : 'dashboard', false);
})();
