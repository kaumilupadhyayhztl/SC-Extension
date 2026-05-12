// ── State ──────────────────────────────────────────────────────────────────
let templates     = [];
let structureList = [];
let currentTpl    = null;
let currentItem   = null;   // item being edited  { _idx, name, isPage, ... }
let mappings      = {};     // fieldId → value
let pickerActive  = false;
let nameMappingMode = false;
let activeFieldId   = null;
let apiUrl = '';
let lastSelected = null;

// ── Init ───────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  const result = await chrome.storage.sync.get(['apiUrl']);
  if (result.apiUrl) {
    apiUrl = result.apiUrl;
    document.getElementById('api-url').value = apiUrl;
  }

  // Receive element picks from content.js
  chrome.runtime.onMessage.addListener(msg => {
    if (msg.type === 'ELEMENT_SELECTED') handleElementSelected(msg.payload);
  });
});

// ── Config ─────────────────────────────────────────────────────────────────
async function saveConfig() {
  apiUrl = document.getElementById('api-url').value.trim().replace(/\/$/, '');
  await chrome.storage.sync.set({ apiUrl });
  showStatus('✅ Config saved!', 'ok');
}

// ── Picker toggle ──────────────────────────────────────────────────────────
async function togglePicker() {
  pickerActive = !pickerActive;
  const btn = document.getElementById('picker-btn');
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    await chrome.tabs.sendMessage(tab.id, {
      type: pickerActive ? 'ACTIVATE_PICKER' : 'DEACTIVATE_PICKER'
    });
    btn.textContent = pickerActive ? '🟢 Picker: ON' : '🔴 Picker: OFF';
    btn.className   = pickerActive ? 'picker-btn on'  : 'picker-btn off';
    if (!pickerActive) { nameMappingMode = false; activeFieldId = null; resetMapBtns(); }
  } catch {
    showStatus('❌ Cannot inject into this page', 'err');
    pickerActive = false;
    btn.textContent = '🔴 Picker: OFF';
    btn.className   = 'picker-btn off';
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

  // Show sticky selected panel
  document.getElementById('selected-panel').style.display = 'block';
  document.getElementById('selected-tag').textContent  = `<${payload.tag.toLowerCase()}>`;
  document.getElementById('selected-info').textContent = content.slice(0, 120) || '(no text)';

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
  if (!fieldId || !lastSelected) return;
  const content = lastSelected.text || lastSelected.src || lastSelected.alt || lastSelected.href || '';
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

// ── Import Excel ────────────────────────────────────────────────────────────
async function importExcel(input) {
  if (!apiUrl) { showStatus('⚙️ Set API URL first', 'err'); return; }
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  showStatus('⏳ Importing...', 'info');
  try {
    const r    = await fetch(`${apiUrl}/api/import-excel`, { method: 'POST', body: fd });
    const data = await r.json();
    if (data.error) throw new Error(data.error);
    templates = data.templates || [];
    renderTemplates();
    showStatus(`✅ ${templates.length} templates loaded!`, 'ok');
  } catch (e) {
    showStatus(`❌ ${e.message}`, 'err');
  }
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
    div.onclick = () => selectTemplate(t);
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
  document.getElementById('ispage-hint').textContent = val ? '📄 Creates a page in Sitecore' : '🗂 Datasource item';
  document.getElementById('ispage-hint').className   = 'hint ' + (val ? 'blue' : '');
  document.getElementById('parent-row').style.display = val ? '' : 'none';
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

// ── Fields ─────────────────────────────────────────────────────────────────
function renderFields() {
  if (!currentTpl) return;
  const list   = document.getElementById('fields-list');
  const mapped = Object.values(mappings).filter(v => v).length;
  const total  = currentTpl.fields.length;
  const pct    = total ? Math.round(mapped / total * 100) : 0;

  document.getElementById('field-count').textContent = `${mapped} / ${total} fields`;
  document.getElementById('field-pct').textContent   = pct + '%';
  document.getElementById('progress-fill').style.width = pct + '%';

  list.innerHTML = '';
  currentTpl.fields.forEach(f => {
    const val = mappings[f.id] || '';
    const div = document.createElement('div');
    div.className = 'field-row' + (val ? ' mapped' : '');
    div.innerHTML = `
      <span class="field-name" title="${f.id}">${f.id}</span>
      <span class="field-type-lbl">${f.type}</span>
      <input class="field-val" type="text" placeholder="value..." value="${esc(val)}"
             oninput="mappings['${esc(f.id)}']=this.value;renderFields();" />
      <button class="field-map-btn ${activeFieldId === f.id ? 'active' : ''}"
              onclick="setActiveField('${esc(f.id)}',this)" title="Pick from page">📌</button>`;
    list.appendChild(div);
  });
}

function setActiveField(fieldId, btn) {
  if (activeFieldId === fieldId) { activeFieldId = null; btn.classList.remove('active'); return; }
  document.querySelectorAll('.field-map-btn').forEach(b => b.classList.remove('active'));
  activeFieldId = fieldId;
  btn.classList.add('active');
  if (!pickerActive) togglePicker();
  showStatus('⚡ Click any element on the page', 'info');
}

// ── Structure ──────────────────────────────────────────────────────────────
function addToStructure() {
  const name = document.getElementById('item-name').value.trim();
  if (!name || !currentTpl) { showStatus('⚠️ Enter item name and select a template', 'err'); return; }
  const isPage   = document.getElementById('tog-yes').classList.contains('yes-on');
  const parentId = document.getElementById('parent-sel').value;
  const fields   = currentTpl.fields.map(f => ({ id: f.id, type: f.type, value: mappings[f.id] || '' }));

  if (currentItem && currentItem._idx !== undefined) {
    structureList[currentItem._idx] = { ...structureList[currentItem._idx], name, template: currentTpl.name, isPage, parentId, fields };
    showStatus('✅ Item updated!', 'ok');
  } else {
    const safeId = `{NEW-${String(structureList.length + 1).padStart(2,'0')}-${name.replace(/\W+/g,'-').toUpperCase().slice(0,12)}}`;
    structureList.push({ id: safeId, name, template: currentTpl.name, isPage, parentId, fields });
    showStatus('✅ Added to structure!', 'ok');
  }

  renderStructureTree();
  // Reset form for next item
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
  const tree = document.getElementById('struct-tree');
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
    div.innerHTML = `
      <div style="min-width:0;">
        <div class="si-name">${item.isPage ? '📄' : '🗂'} ${item.name}</div>
        <div class="si-tpl">${item.template}</div>
      </div>
      <div class="si-actions">
        <button class="si-btn" onclick="editItem(${i})" title="Edit">✏️</button>
        <button class="si-btn" onclick="removeItem(${i})" title="Remove">🗑️</button>
      </div>`;
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
  if (currentItem?._idx === i) { currentItem = null; }
  renderStructureTree();
}

// ── Export Excel ────────────────────────────────────────────────────────────
async function exportExcel() {
  if (!apiUrl) { showStatus('⚙️ Set API URL first', 'err'); return; }
  if (!structureList.length) { showStatus('⚠️ Add at least one item first', 'err'); return; }
  showStatus('⏳ Exporting...', 'info');
  try {
    const r = await fetch(`${apiUrl}/api/export-excel`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ items: structureList })
    });
    if (!r.ok) throw new Error(await r.text());
    const blob = await r.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = 'SitecoreContentExport.xlsx'; a.click();
    URL.revokeObjectURL(url);
    showStatus('✅ Exported successfully!', 'ok');
  } catch (e) { showStatus(`❌ ${e.message}`, 'err'); }
}

// ── Section collapse ────────────────────────────────────────────────────────
function toggleSection(id) {
  const body  = document.getElementById(`${id}-body`);
  const arrow = document.getElementById(`${id}-arrow`);
  const open  = body.style.display !== 'none';
  body.style.display  = open ? 'none' : '';
  if (arrow) arrow.textContent = open ? '▶' : '▼';
}

// ── Status ──────────────────────────────────────────────────────────────────
let _st;
function showStatus(msg, type = 'info') {
  const bar = document.getElementById('status-bar');
  bar.textContent = msg;
  bar.className   = `status-bar ${type}`;
  bar.style.display = 'block';
  clearTimeout(_st);
  _st = setTimeout(() => { bar.style.display = 'none'; }, 3000);
}

// ── Utils ───────────────────────────────────────────────────────────────────
function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/"/g,'&quot;')
    .replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/'/g,'&#39;');
}
