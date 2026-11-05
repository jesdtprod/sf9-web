(function () {
  'use strict';
  const controls = [];
  let active = null;

  function close(restoreFocus) {
    if (!active) return;
    const control = active;
    active = null;
    control.panel.hidden = true;
    control.trigger.setAttribute('aria-expanded', 'false');
    control.trigger.removeAttribute('aria-activedescendant');
    control.field.classList.remove('dropdown-open');
    if (restoreFocus) control.trigger.focus();
  }

  function position(control) {
    if (active !== control) return;
    const rect = control.field.getBoundingClientRect();
    if (!rect.width || !rect.height) { close(false); return; }
    const viewport = window.visualViewport;
    const leftEdge = viewport ? viewport.offsetLeft : 0;
    const topEdge = viewport ? viewport.offsetTop : 0;
    const width = viewport ? viewport.width : window.innerWidth;
    const height = viewport ? viewport.height : window.innerHeight;
    const margin = 8, gap = 6;
    const panelWidth = Math.min(Math.max(rect.width, control.select.id === 'selStudent' ? 300 : 230), width - margin * 2);
    const left = Math.max(leftEdge + margin, Math.min(rect.left, leftEdge + width - panelWidth - margin));
    const below = topEdge + height - rect.bottom - gap - margin;
    const above = rect.top - topEdge - gap - margin;
    const placeAbove = below < 190 && above > below;
    const available = Math.max(90, placeAbove ? above : below);
    control.panel.style.width = panelWidth + 'px';
    control.panel.style.left = left + 'px';
    control.panel.style.maxHeight = Math.min(340, available) + 'px';
    control.list.style.maxHeight = Math.max(40, Math.min(280, available - 60)) + 'px';
    control.panel.style.top = (placeAbove ? Math.max(topEdge + margin, rect.top - gap - control.panel.offsetHeight) : rect.bottom + gap) + 'px';
  }

  function refresh(control) {
    const option = control.select.selectedOptions[0];
    control.label.textContent = option ? option.textContent : (control.select.id === 'selStudent' ? 'Select a student' : 'Select a view');
    control.trigger.title = control.label.textContent;
    control.trigger.disabled = control.select.disabled || !control.select.options.length;
    if (active === control) render(control);
  }

  function choose(control, option) {
    if (!option || option.disabled) return;
    const changed = control.select.value !== option.value;
    control.select.value = option.value;
    if (changed) control.select.dispatchEvent(new Event('change', { bubbles: true }));
    refresh(control);
    close(true);
  }

  function highlight(control, index) {
    const options = [...control.list.querySelectorAll('[role="option"]')];
    if (!options.length) return;
    control.highlighted = Math.max(0, Math.min(index, options.length - 1));
    options.forEach((option, i) => option.classList.toggle('is-active', i === control.highlighted));
    const current = options[control.highlighted];
    control.search.setAttribute('aria-activedescendant', current.id);
    current.scrollIntoView({ block: 'nearest' });
  }

  function render(control) {
    const query = control.search.value.trim().toLocaleLowerCase();
    control.list.replaceChildren();
    const matches = [...control.select.options].filter(option => option.textContent.toLocaleLowerCase().includes(query));
    matches.forEach((option, index) => {
      const item = document.createElement('div');
      item.id = control.list.id + '-option-' + index;
      item.className = 'dropdown-option';
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', String(option.selected));
      item.textContent = option.textContent;
      if (option.selected) {
        const check = document.createElement('span');
        check.className = 'dropdown-check';
        check.setAttribute('aria-hidden', 'true');
        check.textContent = '\u2713';
        item.appendChild(check);
      }
      item.addEventListener('click', () => choose(control, option));
      control.list.appendChild(item);
    });
    control.empty.hidden = matches.length > 0;
    control.highlighted = -1;
    control.search.removeAttribute('aria-activedescendant');
    position(control);
  }

  function open(control) {
    if (control.trigger.disabled) return;
    if (active === control) { close(true); return; }
    close(false);
    active = control;
    control.search.value = '';
    control.panel.hidden = false;
    control.field.classList.add('dropdown-open');
    control.trigger.setAttribute('aria-expanded', 'true');
    render(control);
    control.search.focus();
  }

  function makeDropdown(id) {
    const select = document.getElementById(id), field = select.parentElement;
    const name = select.getAttribute('aria-label');
    select.classList.add('dropdown-native');
    select.tabIndex = -1;
    select.setAttribute('aria-hidden', 'true');
    const trigger = document.createElement('button');
    trigger.type = 'button';
    trigger.className = 'dropdown-trigger';
    trigger.setAttribute('role', 'combobox');
    trigger.setAttribute('aria-label', name);
    trigger.setAttribute('aria-expanded', 'false');
    trigger.setAttribute('aria-haspopup', 'listbox');
    trigger.innerHTML = '<span class="dropdown-trigger-label"></span><svg viewBox="0 0 24 24" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg>';
    field.appendChild(trigger);
    const panel = document.createElement('div');
    panel.className = 'dropdown-panel no-print';
    panel.hidden = true;
    panel.innerHTML = '<div class="dropdown-search-wrap"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m16 16 4 4"/></svg><input class="dropdown-search" type="search" autocomplete="off" spellcheck="false"></div><div class="dropdown-list" role="listbox"></div><div class="dropdown-empty" hidden>No results found</div>';
    document.body.appendChild(panel);
    const search = panel.querySelector('input'), list = panel.querySelector('.dropdown-list');
    list.id = id + '-listbox';
    list.setAttribute('aria-label', name);
    search.setAttribute('aria-label', 'Search ' + name.toLowerCase());
    search.setAttribute('role', 'combobox');
    search.setAttribute('aria-autocomplete', 'list');
    search.setAttribute('aria-expanded', 'true');
    search.setAttribute('aria-controls', list.id);
    trigger.setAttribute('aria-controls', list.id);
    const control = { select, field, trigger, label: trigger.querySelector('.dropdown-trigger-label'), panel, search, list, empty: panel.querySelector('.dropdown-empty'), highlighted: -1 };
    trigger.addEventListener('click', () => open(control));
    trigger.addEventListener('keydown', event => { if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { event.preventDefault(); open(control); highlight(control, event.key === 'ArrowDown' ? 0 : list.children.length - 1); } });
    search.addEventListener('input', () => render(control));
    search.addEventListener('keydown', event => {
      if (event.key === 'Escape') { event.preventDefault(); close(true); }
      else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { event.preventDefault(); highlight(control, control.highlighted + (event.key === 'ArrowDown' ? 1 : -1)); }
      else if (event.key === 'Enter') {
        event.preventDefault();
        const matches = [...select.options].filter(option => option.textContent.toLocaleLowerCase().includes(search.value.trim().toLocaleLowerCase()));
        choose(control, matches[control.highlighted < 0 ? 0 : control.highlighted]);
      } else if (event.key === 'Tab') close(false);
    });
    select.addEventListener('change', () => refresh(control));
    new MutationObserver(() => refresh(control)).observe(select, { childList: true, subtree: true, attributes: true });
    controls.push(control);
    refresh(control);
  }

  makeDropdown('selFormat');
  makeDropdown('selCardView');
  makeDropdown('selStudent');
  document.addEventListener('pointerdown', event => { if (active && !active.panel.contains(event.target) && !active.field.contains(event.target)) close(false); });
  document.addEventListener('keydown', event => { if (active && event.key === 'Escape') { event.preventDefault(); close(true); } });
  window.addEventListener('resize', () => { if (active) position(active); });
  window.addEventListener('scroll', () => { if (active) position(active); }, true);
  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', () => { if (active) position(active); });
    window.visualViewport.addEventListener('scroll', () => { if (active) position(active); });
  }
  window.ECardDropdowns = { sync() { controls.forEach(refresh); if (active && (!active.field.getClientRects().length || active.trigger.disabled)) close(false); } };
})();
