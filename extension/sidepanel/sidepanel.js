// ── State ──────────────────────────────────────────────────────────────────
let templates     = [];
let structureList = [];
let currentTpl    = null;
let currentItem   = null;   // item being edited  { _idx, name, isPage, ... }
let mappings      = {};     // fieldId → value
let pickerActive  = false;
let pickerTabId   = null;   // tab that has the picker active — used to deactivate reliably
let nameMappingMode = false;
let activeFieldId   = null;
let lastSelected    = null;

// ── Init ───────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {

  // ── Wire all event listeners (MV3 forbids inline onclick/onchange) ──
  document.getElementById('picker-btn')        .addEventListener('click',  togglePicker);
  document.getElementById('excel-input')       .addEventListener('change', e => importExcel(e.target));
  document.getElementById('export-btn')        .addEventListener('click',  exportExcel);
  document.getElementById('tpl-search')        .addEventListener('input',  e => filterTpl(e.target.value));
  document.getElementById('name-map-btn')      .addEventListener('click',  toggleNameMapping);
  document.getElementById('tog-yes')           .addEventListener('click',  () => setIsPage(true));
  document.getElementById('tog-no')            .addEventListener('click',  () => setIsPage(false));
  document.getElementById('save-fields-btn')   .addEventListener('click',  saveFields);
  document.getElementById('add-structure-btn') .addEventListener('click',  addToStructure);
  document.getElementById('map-field-btn')     .addEventListener('click',  mapSelectedToField);
  document.getElementById('close-selected-btn').addEventListener('click',  closeSelectedPanel);

  // Section collapse toggles
  document.getElementById('hdr-templates') .addEventListener('click', () => toggleSection('templates'));
  document.getElementById('hdr-structure') .addEventListener('click', () => toggleSection('structure'));
  document.getElementById('hdr-item')      .addEventListener('click', () => toggleSection('item'));

  // Receive element picks from content.js + tab-load events from background.js
  chrome.runtime.onMessage.addListener(msg => {
    if (msg.type === 'ELEMENT_SELECTED') handleElementSelected(msg.payload);
    if (msg.type === 'TAB_LOADED')       onTabLoaded(msg.tabId);
  });
});

// ── Picker toggle ──────────────────────────────────────────────────────────
async function togglePicker() {
  const btn = document.getElementById('picker-btn');

  // ── DISABLE SELECTION ──
  if (pickerActive) {
    pickerActive    = false;
    nameMappingMode = false;
    activeFieldId   = null;
    resetMapBtns();
    setPickerBtn(false);

    if (pickerTabId !== null) {
      try {
        await chrome.tabs.sendMessage(pickerTabId, { type: 'DEACTIVATE_PICKER' });
      } catch {
        // Tab navigated or closed — content script already gone, fine
      }
      pickerTabId = null;
    }
    showStatus('Selection mode disabled', 'info');
    return;
  }

  // ── ENABLE SELECTION ──
  try {
    const tabs = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    const tab  = tabs[0];

    if (!tab || !tab.id) throw new Error('No active tab found');
    if (!tab.url || tab.url.startsWith('chrome://') || tab.url.startsWith('chrome-extension://') || tab.url.startsWith('edge://')) {
      throw new Error('Cannot run on browser system pages. Please navigate to a normal website.');
    }

    await activatePickerOnTab(tab.id);
    pickerActive = true;
    pickerTabId  = tab.id;
    setPickerBtn(true);
    showStatus('✅ Selection mode ON — click any element on the page', 'info');

  } catch (e) {
    console.error('[SC Tool] Selection error:', e);
    showStatus(`❌ ${e.message}`, 'err', 6000);
    pickerActive = false;
    pickerTabId  = null;
    setPickerBtn(false);
  }
}

// Inject content script if needed, then send ACTIVATE_PICKER
async function activatePickerOnTab(tabId) {
  try {
    await chrome.tabs.sendMessage(tabId, { type: 'ACTIVATE_PICKER' });
  } catch {
    console.log('[SC Tool] Injecting content script into tab', tabId);
    await chrome.scripting.executeScript({ target: { tabId }, files: ['content.js'] });
    await new Promise(r => setTimeout(r, 200));
    await chrome.tabs.sendMessage(tabId, { type: 'ACTIVATE_PICKER' });
  }
}

// Update the picker button appearance
function setPickerBtn(on) {
  const btn = document.getElementById('picker-btn');
  if (on) {
    btn.textContent = '✅ Selecting… Click to Stop';
    btn.className   = 'picker-btn on';
  } else {
    btn.textContent = '🎯 Enable Selection';
    btn.className   = 'picker-btn off';
  }
}

// Re-activate picker when user navigates to a new page while selection mode is ON
async function onTabLoaded(tabId) {
  if (!pickerActive || tabId !== pickerTabId) return;
  try {
    await new Promise(r => setTimeout(r, 400));
    await activatePickerOnTab(tabId);
    showStatus('✅ Selection mode resumed on new page — click any element', 'info');
  } catch (e) {
    console.log('[SC Tool] Could not resume selection after navigation:', e);
  }
}

function resetMapBtns() {
  document.getElementById('name-map-btn')?.classList.remove('active');
  document.querySelectorAll('.field-map-btn').forEach(b => b.classList.remove('active'));
}

// ── Element Selected ────────────────────────────────────────────────────────
function handleElementSelected(payload) {
  lastSelected = payload;
  const content = payload.text || payload.src || payload.alt || payload.href || '';

  document.getElementById('selected-panel').style.display = 'block';
  document.getElementById('selected-tag').textContent = `<${payload.tag.toLowerCase()}>`;
  document.getElementById('selected-info').value = content || '(no text)';

  // Populate field dropdown
  const sel = document.getElementById('map-field-sel');
  sel.innerHTML = '<option value="">— map to field —</option>';
  if (currentTpl) {
    currentTpl.fields.forEach(f => {
      const o = document.createElement('option');
      o.value = f.id; o.textContent = f.id;
      if (f.id === activeFieldId) o.selected = true;
      sel.appendChild(o);
    });
  }

  // Auto-fill name if name-mapping mode is on
  if (nameMappingMode) {
    document.getElementById('item-name').value = payload.text.slice(0, 100);
    nameMappingMode = false;
    document.getElementById('name-map-btn').classList.remove('active');
    document.getElementById('name-map-hint').style.display = 'none';
    showStatus('✅ Name captured!', 'ok');
    return;
  }

  // Auto-fill if a specific field's map button was active
  if (activeFieldId) {
    mappings[activeFieldId] = content;
    renderFields();
    showStatus(`✅ Mapped → ${activeFieldId}`, 'ok');
    activeFieldId = null;
    resetMapBtns();
  }
}

function mapSelectedToField() {
  const fieldId = document.getElementById('map-field-sel').value;
  if (!fieldId) return;
  const content = document.getElementById('selected-info').value || '';
  mappings[fieldId] = content;
  renderFields();
  closeSelectedPanel();
  showStatus(`✅ Mapped → ${fieldId}`, 'ok');
}

function closeSelectedPanel() {
  document.getElementById('selected-panel').style.display = 'none';
  lastSelected = null;
}

// ── Name mapping ────────────────────────────────────────────────────────────
function toggleNameMapping() {
  nameMappingMode = !nameMappingMode;
  document.getElementById('name-map-btn').classList.toggle('active', nameMappingMode);
  document.getElementById('name-map-hint').style.display = nameMappingMode ? 'block' : 'none';
  if (nameMappingMode && !pickerActive) togglePicker();
}

// ── Import Excel (client-side via SheetJS) ──────────────────────────────────
function importExcel(input) {
  const file = input.files[0];
  if (!file) return;
  showStatus('⏳ Reading file…', 'info', 30000);

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });

      // ── TemplateFields sheet → templates ──
      const tfSheet = wb.Sheets['TemplateFields'];
      if (!tfSheet) throw new Error('"TemplateFields" sheet not found in the Excel file');

      const tfRows = XLSX.utils.sheet_to_json(tfSheet, { header: 1 });
      const parsed = [];
      tfRows.slice(1).forEach(row => {
        const tplName = String(row[0] || '').trim();
        if (!tplName) return;
        const fields = [];
        for (let i = 1; i < row.length; i += 2) {
          const fname = String(row[i]   || '').trim();
          const ftype = String(row[i+1] || '').trim();
          if (!fname) continue;
          fields.push({ id: fname, type: ftype || 'Single-Line Text', currentValue: '' });
        }
        parsed.push({ name: tplName, fields });
      });

      if (!parsed.length) throw new Error('No templates found — check your TemplateFields sheet');

      templates = parsed;
      renderTemplates();
      showStatus(`✅ ${templates.length} template${templates.length > 1 ? 's' : ''} loaded!`, 'ok', 5000);
      console.log('[SC Tool] Templates loaded:', templates.length);

    } catch (err) {
      console.error('[SC Tool] Import error:', err);
      showStatus(`❌ ${err.message}`, 'err', 10000);
    }
  };
  reader.onerror = () => showStatus('❌ Could not read the file', 'err', 8000);
  reader.readAsArrayBuffer(file);
  input.value = '';
}

// ── Templates ──────────────────────────────────────────────────────────────
function renderTemplates(filter = '') {
  const list     = document.getElementById('tpl-list');
  const badge    = document.getElementById('tpl-count');
  const q        = filter.toLowerCase();
  const filtered = templates.filter(t => t.name.toLowerCase().includes(q));

  badge.textContent   = templates.length;
  badge.style.display = templates.length ? '' : 'none';
  list.innerHTML = '';

  if (!filtered.length) {
    list.innerHTML = `<div class="empty-msg">${templates.length ? 'No match' : '📥 Import Excel to load templates'}</div>`;
    return;
  }

  filtered.forEach(t => {
    const div = document.createElement('div');
    div.className = 'tpl-item' + (currentTpl?.name === t.name ? ' active' : '');
    div.innerHTML = `<span class="tpl-name">${t.name}</span><span class="tpl-fields">${t.fields.length}f</span>`;
    div.addEventListener('click', () => selectTemplate(t));
    list.appendChild(div);
  });
}

function filterTpl(val) { renderTemplates(val); }

function selectTemplate(tpl) {
  currentTpl  = tpl;
  currentItem = null;
  mappings    = {};
  renderTemplates(document.getElementById('tpl-search').value);
  showItemSection(tpl);
}

// ── Item section ────────────────────────────────────────────────────────────
function showItemSection(tpl) {
  document.getElementById('item-section').style.display = '';
  document.getElementById('item-body').style.display    = '';
  document.getElementById('item-tpl-lbl').textContent   = tpl.name;
  document.getElementById('item-name').value = currentItem?.name || '';
  setIsPage(currentItem ? currentItem.isPage : true);
  renderParentSelect();
  renderFields();
}

function setIsPage(val) {
  document.getElementById('tog-yes').className = 'tog' + (val  ? ' yes-on' : '');
  document.getElementById('tog-no').className  = 'tog' + (!val ? ' no-on'  : '');
  document.getElementById('ispage-hint').textContent = val
    ? '📄 Creates a page in Sitecore'
    : '🗂 Datasource — select which page it belongs to';
  document.getElementById('ispage-hint').className = 'hint ' + (val ? 'blue' : 'orange');
  // Parent dropdown shows for Datasource (No), hidden for Page (Yes)
  document.getElementById('parent-row').style.display = !val ? '' : 'none';
  if (!currentItem) currentItem = { isPage: val };
  else currentItem.isPage = val;
}

function renderParentSelect() {
  const sel   = document.getElementById('parent-sel');
  const pages = structureList.filter(i => i.isPage);
  sel.innerHTML = '<option value="">— Home page (root) —</option>';
  pages.forEach(p => {
    const o = document.createElement('option');
    o.value = p.id; o.textContent = p.name;
    if (currentItem?.parentId === p.id) o.selected = true;
    sel.appendChild(o);
  });
}

// ── Fields progress (update counters without re-rendering inputs) ───────────
function updateFieldProgress() {
  if (!currentTpl) return;
  const mapped = Object.values(mappings).filter(v => v).length;
  const total  = currentTpl.fields.length;
  const pct    = total ? Math.round(mapped / total * 100) : 0;
  document.getElementById('field-count').textContent   = `${mapped} / ${total} fields`;
  document.getElementById('field-pct').textContent     = pct + '%';
  document.getElementById('progress-fill').style.width = pct + '%';
  document.querySelectorAll('#fields-list .field-row').forEach((row, i) => {
    const f = currentTpl.fields[i];
    if (f) row.className = 'field-row' + (mappings[f.id] ? ' mapped' : '');
  });
}

// ── Fields (full render — only on template change, not on typing) ───────────
function renderFields() {
  if (!currentTpl) return;
  const list = document.getElementById('fields-list');

  updateFieldProgress();
  list.innerHTML = '';

  currentTpl.fields.forEach(f => {
    const val = mappings[f.id] || '';
    const div = document.createElement('div');
    div.className = 'field-row' + (val ? ' mapped' : '');

    const nameSpan = document.createElement('span');
    nameSpan.className = 'field-name';
    nameSpan.title = f.id;
    nameSpan.textContent = f.id;

    const typeSpan = document.createElement('span');
    typeSpan.className = 'field-type-lbl';
    typeSpan.textContent = f.type;

    const inp = document.createElement('input');
    inp.className   = 'field-val';
    inp.type        = 'text';
    inp.placeholder = 'value…';
    inp.value       = val;
    // Only update mapping + progress — do NOT call renderFields() (would lose focus)
    inp.addEventListener('input', () => { mappings[f.id] = inp.value; updateFieldProgress(); });

    const btn = document.createElement('button');
    btn.className = 'field-map-btn' + (activeFieldId === f.id ? ' active' : '');
    btn.title     = 'Pick from page';
    btn.textContent = '📌';
    btn.addEventListener('click', () => setActiveField(f.id, btn));

    div.appendChild(nameSpan);
    div.appendChild(typeSpan);
    div.appendChild(inp);
    div.appendChild(btn);
    list.appendChild(div);
  });
}

function setActiveField(fieldId, btn) {
  if (activeFieldId === fieldId) {
    activeFieldId = null;
    btn.classList.remove('active');
    return;
  }
  document.querySelectorAll('.field-map-btn').forEach(b => b.classList.remove('active'));
  activeFieldId = fieldId;
  btn.classList.add('active');
  if (!pickerActive) togglePicker();
  showStatus('⚡ Click any element on the page', 'info');
}

// ── Structure ──────────────────────────────────────────────────────────────
function addToStructure() {
  const name = document.getElementById('item-name').value.trim();
  if (!name || !currentTpl) {
    showStatus('⚠️ Enter item name and select a template', 'err');
    return;
  }
  const isPage   = document.getElementById('tog-yes').classList.contains('yes-on');
  const parentId = document.getElementById('parent-sel').value;
  const fields   = currentTpl.fields.map(f => ({ id: f.id, type: f.type, value: mappings[f.id] || '' }));

  if (currentItem && currentItem._idx !== undefined) {
    structureList[currentItem._idx] = {
      ...structureList[currentItem._idx], name, template: currentTpl.name, isPage, parentId, fields
    };
    showStatus('✅ Item updated!', 'ok');
  } else {
    const safeId = `{NEW-${String(structureList.length + 1).padStart(2,'0')}-${name.replace(/\W+/g,'-').toUpperCase().slice(0,12)}}`;
    structureList.push({ id: safeId, name, template: currentTpl.name, isPage, parentId, fields });
    showStatus('✅ Added to structure!', 'ok');
  }

  renderStructureTree();
  document.getElementById('item-name').value = '';
  mappings    = {};
  currentItem = null;
  renderFields();
}

function saveFields() {
  if (!currentItem || currentItem._idx === undefined) { addToStructure(); return; }
  structureList[currentItem._idx].fields =
    currentTpl.fields.map(f => ({ id: f.id, type: f.type, value: mappings[f.id] || '' }));
  renderStructureTree();
  showStatus('✅ Fields saved!', 'ok');
}

function renderStructureTree() {
  const tree  = document.getElementById('struct-tree');
  const badge = document.getElementById('struct-count');

  badge.textContent   = structureList.length;
  badge.style.display = structureList.length ? '' : 'none';

  if (!structureList.length) {
    tree.innerHTML = '<div class="empty-msg">Select a template above and add items here.</div>';
    return;
  }
  tree.innerHTML = '';
  structureList.forEach((item, i) => {
    const div = document.createElement('div');
    div.className = 'struct-item' + (currentItem?._idx === i ? ' active' : '');

    const info = document.createElement('div');
    info.style.minWidth = '0';
    info.innerHTML = `<div class="si-name">${item.isPage ? '📄' : '🗂'} ${item.name}</div>
                      <div class="si-tpl">${item.template}</div>`;

    const actions = document.createElement('div');
    actions.className = 'si-actions';

    const editBtn = document.createElement('button');
    editBtn.className   = 'si-btn';
    editBtn.title       = 'Edit';
    editBtn.textContent = '✏️';
    editBtn.addEventListener('click', e => { e.stopPropagation(); editItem(i); });

    const delBtn = document.createElement('button');
    delBtn.className   = 'si-btn';
    delBtn.title       = 'Remove';
    delBtn.textContent = '🗑️';
    delBtn.addEventListener('click', e => { e.stopPropagation(); removeItem(i); });

    actions.appendChild(editBtn);
    actions.appendChild(delBtn);
    div.appendChild(info);
    div.appendChild(actions);
    tree.appendChild(div);
  });
}

function editItem(i) {
  const item = structureList[i];
  const tpl  = templates.find(t => t.name === item.template);
  if (!tpl) { showStatus('⚠️ Template not found', 'err'); return; }
  currentTpl  = tpl;
  currentItem = { ...item, _idx: i };
  mappings    = {};
  (item.fields || []).forEach(f => { mappings[f.id] = f.value || ''; });
  renderTemplates(document.getElementById('tpl-search').value);
  showItemSection(tpl);
  document.getElementById('item-name').value = item.name;
  setIsPage(item.isPage);
  renderParentSelect();
  renderFields();
}

function removeItem(i) {
  if (!confirm(`Remove "${structureList[i].name}"?`)) return;
  structureList.splice(i, 1);
  if (currentItem?._idx === i) currentItem = null;
  renderStructureTree();
}

// ── Export Excel (client-side via SheetJS) ──────────────────────────────────
function exportExcel() {
  if (!structureList.length) {
    showStatus('⚠️ Add at least one item to the structure first', 'err');
    return;
  }

  try {
    const wb = XLSX.utils.book_new();

    // ── Items sheet ──
    const itemHeader = ['ID','ParentID','Name','Template','Path','IsPage','DatasourceIDs','SortOrder','Language','Version','WorkflowState','Publish'];
    const itemRows   = [itemHeader];
    structureList.forEach((item, i) => {
      itemRows.push([
        item.id       || `{ITEM-${i+1}-ID}`,
        item.parentId || '',
        item.name     || item.template,
        item.template,
        item.path     || '',
        item.isPage   ? 'Yes' : 'No',
        item.datasourceId || '',
        (i + 1) * 100,
        'en', 1, 'Draft', 'Yes'
      ]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(itemRows), 'Items');

    // ── Fields sheet ──
    const maxFields  = Math.max(...structureList.map(it => (it.fields || []).length), 1);
    const fieldHeader = ['ItemID', 'Language', 'Version'];
    for (let i = 0; i < maxFields; i++) fieldHeader.push('FieldName', 'FieldValue', 'FieldType');
    const fieldRows = [fieldHeader];
    structureList.forEach((item, i) => {
      const row = [item.id || `{ITEM-${i+1}-ID}`, 'en', 1];
      (item.fields || []).forEach(f => row.push(f.id, f.value || '', f.type));
      while (row.length < 3 + maxFields * 3) row.push('', '', '');
      fieldRows.push(row);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fieldRows), 'Fields');

    // ── Media + Links placeholder sheets ──
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['MediaID','FileName','FilePath','DestinationFolder','Alt','Title','Language']
    ]), 'Media');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['ItemID','FieldName','LinkType','LinkText','Url','Target','Anchor','Class']
    ]), 'Links');

    // ── Trigger browser download ──
    XLSX.writeFile(wb, 'SitecoreContentExport.xlsx');
    showStatus('✅ Exported successfully!', 'ok');

  } catch (err) {
    console.error('[SC Tool] Export error:', err);
    showStatus(`❌ Export failed: ${err.message}`, 'err');
  }
}

// ── Section collapse ────────────────────────────────────────────────────────
function toggleSection(id) {
  const body  = document.getElementById(`${id}-body`);
  const arrow = document.getElementById(`${id}-arrow`);
  const open  = body.style.display !== 'none';
  body.style.display = open ? 'none' : '';
  if (arrow) arrow.textContent = open ? '▶' : '▼';
}

// ── Status bar ──────────────────────────────────────────────────────────────
let _statusTimer;
function showStatus(msg, type = 'info', duration = 3000) {
  const bar = document.getElementById('status-bar');
  bar.textContent   = msg;
  bar.className     = `status-bar ${type}`;
  bar.style.display = 'block';
  clearTimeout(_statusTimer);
  if (duration > 0) _statusTimer = setTimeout(() => { bar.style.display = 'none'; }, duration);
}
