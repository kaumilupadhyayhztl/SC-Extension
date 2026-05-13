let pickerActive = false;
let lastEl = null;

// Inject highlight styles into the page once
(function injectStyles() {
  if (document.getElementById('__sct_style')) return;
  const s = document.createElement('style');
  s.id = '__sct_style';
  s.textContent = `
    .__sct-hover    { outline: 2px solid #378ADD !important; cursor: crosshair !important; }
    .__sct-selected { outline: 3px solid #E8A020 !important; background: rgba(232,160,32,.08) !important; }
  `;
  (document.head || document.documentElement).appendChild(s);
})();

// Listen for commands from the side panel
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.type === 'ACTIVATE_PICKER') {
    pickerActive = true;
    document.body.style.cursor = 'crosshair';
    sendResponse({ ok: true });
  } else if (msg.type === 'DEACTIVATE_PICKER') {
    pickerActive = false;
    document.body.style.cursor = '';
    clearAll();
    sendResponse({ ok: true });
  }
  return true;
});

function clearAll() {
  document.querySelectorAll('.__sct-hover, .__sct-selected').forEach(el => {
    el.classList.remove('__sct-hover', '__sct-selected');
  });
}

document.addEventListener('mouseover', e => {
  if (!pickerActive) return;
  if (lastEl && !lastEl.classList.contains('__sct-selected'))
    lastEl.classList.remove('__sct-hover');
  if (!e.target.classList.contains('__sct-selected')) {
    e.target.classList.add('__sct-hover');
    lastEl = e.target;
  }
}, true);

document.addEventListener('mouseout', e => {
  if (!pickerActive) return;
  if (!e.target.classList.contains('__sct-selected'))
    e.target.classList.remove('__sct-hover');
}, true);

document.addEventListener('click', e => {
  if (!pickerActive) return;
  e.preventDefault();
  e.stopPropagation();
  e.stopImmediatePropagation();

  clearAll();
  e.target.classList.add('__sct-selected');

  // Auto-deactivate picker so the page returns to normal immediately
  pickerActive = false;
  document.body.style.cursor = '';
  // Brief delay so the orange outline is visible, then clear it
  setTimeout(() => {
    e.target.classList.remove('__sct-selected');
  }, 800);

  // Send captured element data to the side panel
  chrome.runtime.sendMessage({
    type: 'ELEMENT_SELECTED',
    payload: {
      tag:  e.target.tagName,
      text: (e.target.innerText || '').trim().slice(0, 500),
      src:  e.target.src  || e.target.getAttribute('data-src') || '',
      alt:  e.target.alt  || '',
      href: e.target.href || ''
    }
  });

  // Tell the side panel to sync its picker button to OFF
  chrome.runtime.sendMessage({ type: 'PICKER_AUTO_OFF' });
}, true);
