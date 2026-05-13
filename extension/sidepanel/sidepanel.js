// ══════════════════════════════════════════════════════════════
// STATE
// ══════════════════════════════════════════════════════════════
let templates       = [];       // Excel-imported templates (offline fallback)
let structureList   = [];       // offline: items added locally
let currentTpl      = null;     // template currently in the Item form
let currentItem     = null;     // item being edited { _idx, name, isPage, … }
let mappings        = {};       // fieldId → value

let pickerActive    = false;
let pickerTabId     = null;
let nameMappingMode = false;
let activeFieldId   = null;
let lastSelected    = null;

// ── Sitecore live-mode state ──────────────────────────────────
let scConfig        = null;
/*
  scConfig = {
    platform : 'xmcloud' | 'traditional',
    cmUrl    : 'https://…',
    apiKey   : '{GUID}',          // traditional
    token    : 'eyJ…',            // xmcloud Bearer
    tokenExp : timestamp,
    apiMode  : 'graphql' | 'ssc'  // detected on connect
  }
*/
let treeMap         = new Map(); // nodeId → TreeNode
let treeRootId      = null;
let selectedParentNode = null;   // node the user picked for "Create item here"
let scTemplates     = [];        // templates fetched from Sitecore
let isLiveMode      = false;     // true when connected to Sitecore

// ══════════════════════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', async () => {

  // Restore saved connection settings (URL / API key — no secrets)
  const stored = await chrome.storage.local.get(['scSavedConfig']);
  if (stored.scSavedConfig) {
    const s = stored.scSavedConfig;
    document.getElementById('sc-platform').value  = s.platform  || 'xmcloud';
    document.getElementById('sc-cm-url').value    = s.cmUrl     || '';
    document.getElementById('sc-api-key').value   = s.apiKey    || '';
    document.getElementById('sc-root-path').value = s.rootPath  || '/sitecore/content';
    updateCredFields();
  }

  // ── Wire event listeners ──
  document.getElementById('picker-btn')        .addEventListener('click',  togglePicker);
  document.getElementById('excel-input')       .addEventListener('change', e => importExcel(e.target));
  document.getElementById('export-btn')        .addEventListener('click',  exportExcel);
  document.getElementById('tpl-search')        .addEventListener('input',  e => filterTpl(e.target.value));
  document.getElementById('name-map-btn')      .addEventListener('click',  toggleNameMapping);
  document.getElementById('tog-yes')           .addEventListener('click',  () => setIsPage(true));
  document.getElementById('tog-no')            .addEventListener('click',  () => setIsPage(false));
  document.getElementById('save-fields-btn')   .addEventListener('click',  saveFields);
  document.getElementById('add-structure-btn') .addEventListener('click',  addToStructure);
  document.getElementById('create-sc-btn')     .addEventListener('click',  createScItem);
  document.getElementById('map-field-btn')     .addEventListener('click',  mapSelectedToField);
  document.getElementById('close-selected-btn').addEventListener('click',  closeSelectedPanel);

  // Connection
  document.getElementById('sc-platform')      .addEventListener('change', updateCredFields);
  document.getElementById('sc-connect-btn')   .addEventListener('click',  connectSitecore);
  document.getElementById('sc-disconnect-btn').addEventListener('click',  disconnectSitecore);
  document.getElementById('tree-refresh-btn') .addEventListener('click',  refreshTree);
  document.getElementById('tree-search')      .addEventListener('input',  e => filterTree(e.target.value));
  document.getElementById('item-tpl-sel')     .addEventListener('change', onTplSelChange);

  // Section toggles
  document.getElementById('hdr-connection').addEventListener('click', () => toggleSection('connection'));
  document.getElementById('hdr-tree')      .addEventListener('click', () => toggleSection('tree'));
  document.getElementById('hdr-templates') .addEventListener('click', () => toggleSection('templates'));
  document.getElementById('hdr-structure') .addEventListener('click', () => toggleSection('structure'));
  document.getElementById('hdr-item')      .addEventListener('click', () => toggleSection('item'));

  // Messages from content.js / background.js
  chrome.runtime.onMessage.addListener(msg => {
    if (msg.type === 'ELEMENT_SELECTED') handleElementSelected(msg.payload);
    if (msg.type === 'TAB_LOADED')       onTabLoaded(msg.tabId);
  });
});

// ══════════════════════════════════════════════════════════════
// PICKER
// ══════════════════════════════════════════════════════════════
async function togglePicker() {
  if (pickerActive) {
    pickerActive = false; nameMappingMode = false; activeFieldId = null;
    resetMapBtns(); setPickerBtn(false);
    if (pickerTabId !== null) {
      try { await chrome.tabs.sendMessage(pickerTabId, { type: 'DEACTIVATE_PICKER' }); } catch {}
      pickerTabId = null;
    }
    showStatus('Selection mode disabled', 'info');
    return;
  }
  try {
    const tabs = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    const tab  = tabs[0];
    if (!tab?.id) throw new Error('No active tab found');
    if (!tab.url || /^(chrome|chrome-extension|edge):\/\//.test(tab.url))
      throw new Error('Cannot run on browser system pages');
    await activatePickerOnTab(tab.id);
    pickerActive = true; pickerTabId = tab.id;
    setPickerBtn(true);
    showStatus('✅ Selection mode ON — click any element on the page', 'info');
  } catch (e) {
    pickerActive = false; pickerTabId = null; setPickerBtn(false);
    showStatus(`❌ ${e.message}`, 'err', 6000);
  }
}

async function activatePickerOnTab(tabId) {
  try {
    await chrome.tabs.sendMessage(tabId, { type: 'ACTIVATE_PICKER' });
  } catch {
    await chrome.scripting.executeScript({ target: { tabId }, files: ['content.js'] });
    await new Promise(r => setTimeout(r, 200));
    await chrome.tabs.sendMessage(tabId, { type: 'ACTIVATE_PICKER' });
  }
}

async function onTabLoaded(tabId) {
  if (!pickerActive || tabId !== pickerTabId) return;
  try {
    await new Promise(r => setTimeout(r, 400));
    await activatePickerOnTab(tabId);
    showStatus('✅ Selection mode resumed on new page', 'info');
  } catch {}
}

function setPickerBtn(on) {
  const btn = document.getElementById('picker-btn');
  btn.textContent = on ? '✅ Selecting… Click to Stop' : '🎯 Enable Selection';
  btn.className   = on ? 'picker-btn on' : 'picker-btn off';
}

function resetMapBtns() {
  document.getElementById('name-map-btn')?.classList.remove('active');
  document.querySelectorAll('.field-map-btn').forEach(b => b.classList.remove('active'));
}

// ══════════════════════════════════════════════════════════════
// ELEMENT SELECTION
// ══════════════════════════════════════════════════════════════
function handleElementSelected(payload) {
  lastSelected = payload;
  const content = payload.text || payload.src || payload.alt || payload.href || '';
  document.getElementById('selected-panel').style.display = 'block';
  document.getElementById('selected-tag').textContent = `<${payload.tag.toLowerCase()}>`;
  document.getElementById('selected-info').value = content || '(no text)';

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

  if (nameMappingMode) {
    document.getElementById('item-name').value = payload.text.slice(0, 100);
    nameMappingMode = false;
    document.getElementById('name-map-btn').classList.remove('active');
    document.getElementById('name-map-hint').style.display = 'none';
    showStatus('✅ Name captured!', 'ok');
    return;
  }

  if (activeFieldId) {
    mappings[activeFieldId] = content;
    renderFields();
    showStatus(`✅ Mapped → ${activeFieldId}`, 'ok');
    activeFieldId = null; resetMapBtns();
  }
}

function mapSelectedToField() {
  const fieldId = document.getElementById('map-field-sel').value;
  if (!fieldId) return;
  mappings[fieldId] = document.getElementById('selected-info').value || '';
  renderFields(); closeSelectedPanel();
  showStatus(`✅ Mapped → ${fieldId}`, 'ok');
}

function closeSelectedPanel() {
  document.getElementById('selected-panel').style.display = 'none';
  lastSelected = null;
}

function toggleNameMapping() {
  nameMappingMode = !nameMappingMode;
  document.getElementById('name-map-btn').classList.toggle('active', nameMappingMode);
  document.getElementById('name-map-hint').style.display = nameMappingMode ? 'block' : 'none';
  if (nameMappingMode && !pickerActive) togglePicker();
}

// ══════════════════════════════════════════════════════════════
// SITECORE CONNECTION
// ══════════════════════════════════════════════════════════════
function updateCredFields() {
  const platform = document.getElementById('sc-platform').value;
  document.getElementById('xmcloud-creds').style.display    = platform === 'xmcloud'     ? '' : 'none';
  document.getElementById('traditional-creds').style.display = platform === 'traditional' ? '' : 'none';
}

async function connectSitecore() {
  const platform  = document.getElementById('sc-platform').value;
  const cmUrl     = document.getElementById('sc-cm-url').value.trim().replace(/\/$/, '');
  const rootPath  = document.getElementById('sc-root-path').value.trim() || '/sitecore/content';

  if (!cmUrl) { showStatus('⚠️ Enter CM Instance URL', 'err'); return; }

  const cfg = { platform, cmUrl };
  setConnDot('connecting');

  try {
    // ── XM Cloud: get OAuth token ──
    if (platform === 'xmcloud') {
      const clientId     = document.getElementById('sc-client-id').value.trim();
      const clientSecret = document.getElementById('sc-client-secret').value.trim();
      if (!clientId || !clientSecret) throw new Error('Enter Client ID and Client Secret');
      showStatus('🔄 Getting XM Cloud access token…', 'info', 30000);
      const tokenData  = await getXMCloudToken(clientId, clientSecret);
      cfg.token        = tokenData.access_token;
      cfg.tokenExp     = Date.now() + tokenData.expires_in * 1000;

    // ── Traditional: API Key ──
    } else {
      const apiKey = document.getElementById('sc-api-key').value.trim();
      if (!apiKey) throw new Error('Enter the Sitecore API Key');
      cfg.apiKey = apiKey;
    }

    // ── Detect GraphQL vs SSC ──
    showStatus('🔄 Testing connection…', 'info', 30000);
    cfg.apiMode = await detectApiMode(cfg);

    scConfig   = cfg;
    isLiveMode = true;

    // Save non-sensitive settings
    await chrome.storage.local.set({ scSavedConfig: { platform, cmUrl, apiKey: cfg.apiKey || '', rootPath } });

    setConnDot('on');
    showConnectedUi();
    showStatus('✅ Connected! Loading content tree…', 'ok', 5000);
    await loadTree(rootPath);

  } catch (e) {
    console.error('[SC Tool] Connect error:', e);
    setConnDot('off');
    showStatus(`❌ ${e.message}`, 'err', 10000);
  }
}

async function getXMCloudToken(clientId, clientSecret) {
  const res = await fetch('https://auth.sitecorecloud.io/oauth/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type:    'client_credentials',
      client_id:     clientId,
      client_secret: clientSecret,
      audience:      'https://api.sitecorecloud.io'
    })
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`XM Cloud auth failed (${res.status}): ${txt.slice(0, 200)}`);
  }
  return res.json();
}

// Returns 'graphql' or 'ssc' depending on what the instance supports
async function detectApiMode(cfg) {
  // 1 — Try GraphQL
  try {
    const result = await gqlQuery('{ item(path: "/sitecore/content", language: "en") { id } }', {}, cfg);
    if (result.data?.item?.id) return 'graphql';
  } catch {}

  // 2 — Try SSC (verify the response is actual JSON, not an HTML login page)
  try {
    const url = buildSscUrl(cfg, `-/item/v1?path=/sitecore/content&database=master`);
    const res  = await fetch(url, { headers: { ...authHeaders(cfg), 'SC_APIKEY': cfg.apiKey || '' } });
    if (res.ok) {
      const ct   = res.headers.get('content-type') || '';
      const text = await res.text();
      if (ct.includes('json') || (text.trim().startsWith('{') || text.trim().startsWith('['))) {
        return 'ssc';
      }
      // Got HTML back — API key likely wrong or SSC not configured
      throw new Error('Sitecore returned an HTML page instead of JSON. Check that your API Key is correct and the Sitecore Services Client (SSC) is enabled on the CM instance.');
    }
    throw new Error(`Sitecore returned HTTP ${res.status}. Check URL and API Key.`);
  } catch (e) {
    if (e.message.includes('HTML') || e.message.includes('HTTP')) throw e;
  }

  throw new Error('Cannot reach Sitecore. Verify the CM URL, API Key, and that the instance is running.');
}

function disconnectSitecore() {
  scConfig   = null;
  isLiveMode = false;
  treeMap.clear();
  treeRootId          = null;
  selectedParentNode  = null;
  scTemplates         = [];
  setConnDot('off');
  showDisconnectedUi();
  showStatus('Disconnected from Sitecore', 'info');
}

function setConnDot(state) {
  const dot = document.getElementById('conn-dot');
  dot.className = `conn-dot ${state}`;
}

function showConnectedUi() {
  document.getElementById('tree-section').style.display        = '';
  document.getElementById('offline-bar').style.display         = 'none';
  document.getElementById('templates-section').style.display   = 'none';
  document.getElementById('structure-section').style.display   = 'none';
  document.getElementById('sc-connect-btn').style.display      = 'none';
  document.getElementById('sc-disconnect-btn').style.display   = '';
  // Collapse connection section to save space
  document.getElementById('connection-body').style.display     = 'none';
  document.getElementById('connection-arrow').textContent      = '▶';
}

function showDisconnectedUi() {
  document.getElementById('tree-section').style.display        = 'none';
  document.getElementById('item-section').style.display        = 'none';
  document.getElementById('offline-bar').style.display         = '';
  document.getElementById('templates-section').style.display   = '';
  document.getElementById('structure-section').style.display   = '';
  document.getElementById('sc-connect-btn').style.display      = '';
  document.getElementById('sc-disconnect-btn').style.display   = 'none';
  document.getElementById('connection-body').style.display     = '';
  document.getElementById('connection-arrow').textContent      = '▼';
}

// ══════════════════════════════════════════════════════════════
// SITECORE API HELPERS
// ══════════════════════════════════════════════════════════════
function authHeaders(cfg) {
  const c = cfg || scConfig;
  if (!c) return {};
  const h = {};
  if (c.platform === 'xmcloud' && c.token) h['Authorization'] = `Bearer ${c.token}`;
  // Pass API key both as header and query param — some Sitecore versions prefer header
  if (c.apiKey) h['SC_APIKEY'] = c.apiKey;
  return h;
}

function buildSscUrl(cfg, endpoint) {
  const c      = cfg || scConfig;
  // Strip curly braces from API key GUID — some Sitecore versions reject them in the URL
  const rawKey = (c.apiKey || '').replace(/[{}]/g, '');
  const key    = rawKey ? `?sc_apikey=${rawKey}` : '';
  return `${c.cmUrl}/sitecore/api/ssc/item/${endpoint}${endpoint.includes('?') && key ? '&' + key.slice(1) : key}`;
}

function buildGqlUrl(cfg) {
  const c = cfg || scConfig;
  const key = c.apiKey ? `?sc_apikey=${c.apiKey}` : '';
  return `${c.cmUrl}/sitecore/api/graph/edge${key}`;
}

async function gqlQuery(query, variables = {}, cfg) {
  const c   = cfg || scConfig;
  const url = buildGqlUrl(c);
  const res = await fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders(c) },
    body:    JSON.stringify({ query, variables })
  });
  if (!res.ok) throw new Error(`GraphQL ${res.status}`);
  return res.json();
}

async function sscGet(endpoint) {
  const url  = buildSscUrl(null, endpoint);
  const res  = await fetch(url, { headers: authHeaders() });
  const text = await res.text();

  // Detect HTML response (login redirect / error page)
  if (text.trim().startsWith('<')) {
    if (!res.ok || res.status === 302 || text.toLowerCase().includes('login')) {
      throw new Error('Sitecore returned a login/HTML page. Verify your API Key is correct and the Sitecore Services Client is enabled.');
    }
    throw new Error('Sitecore returned HTML instead of JSON. Check CM URL and API Key.');
  }

  if (!res.ok) throw new Error(`Sitecore SSC error ${res.status}: ${text.slice(0, 200)}`);

  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`Unexpected response from Sitecore: ${text.slice(0, 100)}`);
  }
}

// ══════════════════════════════════════════════════════════════
// CONTENT TREE
// ══════════════════════════════════════════════════════════════
async function loadTree(rootPath) {
  const container = document.getElementById('tree-container');
  container.innerHTML = '<div class="tree-loading">⏳ Loading tree…</div>';
  try {
    let root;
    if (scConfig.apiMode === 'graphql') {
      root = await fetchNodeGql(rootPath);
    } else {
      root = await fetchRootSsc(rootPath);
    }
    if (!root) throw new Error('Root path not found: ' + rootPath);

    root.depth = 0; root.expanded = false; root.childrenIds = null;
    treeMap.set(root.id, root);
    treeRootId = root.id;

    // Auto-expand root
    await expandTreeNode(root.id);
    renderTree();

  } catch (e) {
    container.innerHTML = `<div class="tree-err">❌ ${e.message}</div>`;
  }
}

async function refreshTree() {
  const rootPath = document.getElementById('sc-root-path').value.trim() || '/sitecore/content';
  treeMap.clear(); treeRootId = null; selectedParentNode = null;
  await loadTree(rootPath);
}

async function expandTreeNode(nodeId) {
  const node = treeMap.get(nodeId);
  if (!node || !node.hasChildren) return;
  if (node.childrenIds !== null) { node.expanded = true; return; } // already loaded

  node.expanded = true;
  node.childrenIds = [];

  try {
    let children;
    if (scConfig.apiMode === 'graphql') {
      children = await fetchChildrenGql(node.path);
    } else {
      children = await fetchChildrenSsc(node.id);
    }

    children.forEach(c => {
      c.depth = node.depth + 1;
      c.expanded = false;
      c.childrenIds = null;
      c.parentId = node.id;
      treeMap.set(c.id, c);
      node.childrenIds.push(c.id);
    });

  } catch (e) {
    node.childrenIds = null;
    console.error('[SC Tool] Tree expand error:', e);
    showStatus('⚠️ Could not load children: ' + e.message, 'err', 5000);
  }
}

async function fetchNodeGql(path) {
  const q = `query($p:String!){item(path:$p,language:"en"){id name path template{id name}hasChildren}}`;
  const r = await gqlQuery(q, { p: path });
  const n = r.data?.item;
  if (!n) return null;
  return { id: n.id, name: n.name, path: n.path, templateId: n.template?.id||'', templateName: n.template?.name||'', hasChildren: !!n.hasChildren };
}

async function fetchChildrenGql(parentPath) {
  const q = `query($p:String!){item(path:$p,language:"en"){children(first:100){results{id name path hasChildren template{id name}}}}}`;
  const r = await gqlQuery(q, { p: parentPath });
  return (r.data?.item?.children?.results || []).map(normaliseGqlNode);
}

function normaliseGqlNode(n) {
  return { id: n.id, name: n.name, path: n.path, templateId: n.template?.id||'', templateName: n.template?.name||'', hasChildren: !!n.hasChildren };
}

async function fetchRootSsc(path) {
  // Do NOT percent-encode slashes — Sitecore expects them as-is in the path param
  const data = await sscGet(`-/item/v1?path=${path}&database=master`);
  return normaliseSscNode(data);
}

async function fetchChildrenSsc(parentId) {
  // Strip curly braces from GUID if present (some Sitecore versions don't accept them in URL)
  const id   = parentId.replace(/[{}]/g, '');
  const data = await sscGet(`${id}/children?database=master`);
  return Array.isArray(data) ? data.map(normaliseSscNode) : [];
}

function normaliseSscNode(n) {
  return {
    id:           n.ItemID   || n.id || '',
    name:         n.ItemName || n.DisplayName || n.name || '',
    path:         n.ItemPath || n.path || '',
    templateId:   n.TemplateID   || n.templateId   || '',
    templateName: n.TemplateName || n.templateName || '',
    hasChildren:  !!(n.HasChildren ?? n.hasChildren)
  };
}

// ── Tree rendering ─────────────────────────────────────────────
function renderTree(filterQ) {
  const container = document.getElementById('tree-container');
  container.innerHTML = '';
  if (!treeRootId) { container.innerHTML = '<div class="empty-msg">No tree loaded</div>'; return; }
  const root = treeMap.get(treeRootId);
  if (!root) return;
  renderTreeNodeEl(root, container, filterQ?.toLowerCase() || '');
}

function renderTreeNodeEl(node, parent, filterQ) {
  const matchSelf = !filterQ || node.name.toLowerCase().includes(filterQ) || node.path.toLowerCase().includes(filterQ);
  const hasMatchingChild = filterQ && getDescendants(node).some(n => n.name.toLowerCase().includes(filterQ));
  if (filterQ && !matchSelf && !hasMatchingChild) return;

  const wrap = document.createElement('div');
  wrap.className = 'tree-node';

  const row = document.createElement('div');
  row.className = 'tree-row' + (selectedParentNode?.id === node.id ? ' active' : '');
  row.style.paddingLeft = (4 + node.depth * 14) + 'px';
  row.dataset.nid = node.id;

  // Toggle
  const tog = document.createElement('span');
  tog.className   = 'tree-toggle';
  tog.textContent = node.hasChildren ? (node.expanded ? '▼' : '▶') : '';
  if (node.hasChildren) {
    tog.addEventListener('click', async e => {
      e.stopPropagation();
      if (node.expanded) {
        node.expanded = false;
      } else {
        tog.textContent = '…';
        await expandTreeNode(node.id);
      }
      renderTree(document.getElementById('tree-search').value);
    });
  }

  // Icon
  const icon = document.createElement('span');
  icon.className   = 'tree-icon';
  icon.textContent = getNodeIcon(node.templateName);

  // Name
  const name = document.createElement('span');
  name.className   = 'tree-name';
  name.textContent = node.name;
  name.title       = node.path;

  // Template
  const meta = document.createElement('span');
  meta.className   = 'tree-meta';
  meta.textContent = node.templateName;

  // Add button
  const addBtn = document.createElement('button');
  addBtn.className   = 'tree-add';
  addBtn.textContent = '+';
  addBtn.title       = `Create item under ${node.path}`;
  addBtn.addEventListener('click', e => { e.stopPropagation(); selectParentNode(node); });

  row.appendChild(tog); row.appendChild(icon); row.appendChild(name);
  row.appendChild(meta); row.appendChild(addBtn);
  row.addEventListener('click', () => selectParentNode(node));
  wrap.appendChild(row);

  // Render children if expanded
  if (node.expanded && node.childrenIds?.length) {
    const childWrap = document.createElement('div');
    node.childrenIds.forEach(cid => {
      const child = treeMap.get(cid);
      if (child) renderTreeNodeEl(child, childWrap, filterQ);
    });
    wrap.appendChild(childWrap);
  }

  parent.appendChild(wrap);
}

function getDescendants(node) {
  const out = [];
  (node.childrenIds || []).forEach(id => {
    const c = treeMap.get(id);
    if (c) { out.push(c); out.push(...getDescendants(c)); }
  });
  return out;
}

function getNodeIcon(templateName) {
  const t = (templateName || '').toLowerCase();
  if (/folder|bucket|node|root/.test(t))    return '📁';
  if (/datasource|rendering|partial/.test(t)) return '🗂';
  return '📄';
}

function filterTree(q) { renderTree(q); }

// ── Select a parent node from the tree (open Item form) ────────
async function selectParentNode(node) {
  selectedParentNode = node;
  renderTree(document.getElementById('tree-search').value); // highlight

  // Load templates from Sitecore
  if (!scTemplates.length) {
    showStatus('⏳ Loading templates from Sitecore…', 'info', 15000);
    await fetchScTemplates();
  }

  openItemForm(node);
}

// ── Fetch templates from Sitecore ─────────────────────────────
async function fetchScTemplates() {
  try {
    let fetched = [];
    if (scConfig.apiMode === 'graphql') {
      fetched = await fetchScTemplatesGql();
    } else {
      fetched = await fetchScTemplatesSsc();
    }
    scTemplates = fetched;
    showStatus(`✅ ${fetched.length} templates loaded from Sitecore`, 'ok', 3000);
  } catch (e) {
    console.error('[SC Tool] Template fetch error:', e);
    showStatus('⚠️ Could not load templates — using Excel templates if available', 'err', 5000);
  }
}

async function fetchScTemplatesGql() {
  // Fetch 2-level deep under /sitecore/templates to find user-defined templates
  const q = `
    query {
      item(path:"/sitecore/templates",language:"en"){
        children(first:50){
          results{
            name
            children(first:100){
              results{
                id name path hasChildren
                children(first:50){
                  results{ id name path template{name}
                    children(first:20){ results{ id name path template{name} } }
                  }
                }
              }
            }
          }
        }
      }
    }
  `;
  const r = await gqlQuery(q);
  const templates = [];
  const sections = r.data?.item?.children?.results || [];
  sections.forEach(section => {
    collectTemplateNodes(section.children?.results || [], templates);
  });
  return templates;
}

function collectTemplateNodes(nodes, out) {
  nodes.forEach(n => {
    // A template node has no further template-type children (leaf or near-leaf)
    const tplName = (n.template?.name || '').toLowerCase();
    if (tplName.includes('template') && !tplName.includes('section')) {
      out.push({ id: n.id, name: n.name, path: n.path, fields: [] });
    } else if (n.children?.results?.length) {
      collectTemplateNodes(n.children.results, out);
    }
  });
}

async function fetchScTemplatesSsc() {
  // Get children of /sitecore/templates recursively (3 levels)
  const root = await sscGet(`-/item/v1?path=/sitecore/templates&database=master`);
  const sections = await fetchChildrenSsc(root.ItemID || root.id);
  const templates = [];
  for (const sec of sections) {
    const children = await fetchChildrenSsc(sec.id);
    for (const c of children) {
      templates.push({ id: c.id, name: c.name, path: c.path, fields: [] });
      if (c.hasChildren) {
        const sub = await fetchChildrenSsc(c.id);
        sub.forEach(s => templates.push({ id: s.id, name: s.name, path: s.path, fields: [] }));
      }
    }
  }
  return templates;
}

// All available templates (Sitecore + Excel combined)
function getAllTemplates() {
  const combined = [...scTemplates];
  templates.forEach(t => {
    if (!combined.find(c => c.name === t.name)) combined.push(t);
  });
  return combined;
}

// ── Open item creation form ────────────────────────────────────
function openItemForm(parentNode) {
  const allTpls = getAllTemplates();

  // Show the section
  document.getElementById('item-section').style.display = '';

  // "Creating under" label
  document.getElementById('create-under-row').style.display = '';
  document.getElementById('create-under-path').textContent  = parentNode.path;

  // Template picker dropdown
  const tplPicker = document.getElementById('tpl-picker-row');
  const tplSel    = document.getElementById('item-tpl-sel');
  tplPicker.style.display = '';
  tplSel.innerHTML = '<option value="">— select template —</option>';
  allTpls.forEach(t => {
    const o = document.createElement('option');
    o.value = t.id || t.name; o.textContent = t.name; o.dataset.tplname = t.name;
    tplSel.appendChild(o);
  });

  // Show "Create in Sitecore" button, hide "Add to Structure"
  document.getElementById('create-sc-btn').style.display     = '';
  document.getElementById('add-structure-btn').style.display = 'none';

  // Reset form
  currentItem = null; mappings = {}; currentTpl = null;
  document.getElementById('item-name').value = '';
  document.getElementById('fields-list').innerHTML = '';
  document.getElementById('field-count').textContent = '0 / 0';
  document.getElementById('field-pct').textContent   = '0%';
  document.getElementById('progress-fill').style.width = '0%';
  document.getElementById('item-tpl-lbl').textContent = '';
  setIsPage(true);

  // Scroll to item section
  document.getElementById('item-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function onTplSelChange() {
  const sel     = document.getElementById('item-tpl-sel');
  const opt     = sel.options[sel.selectedIndex];
  const tplName = opt?.dataset?.tplname || '';
  const allTpls = getAllTemplates();
  const tpl     = allTpls.find(t => t.name === tplName);
  if (!tpl) return;

  currentTpl = tpl;
  document.getElementById('item-tpl-lbl').textContent = tpl.name;
  mappings = {};
  renderFields();
}

// ── Create item in Sitecore ────────────────────────────────────
async function createScItem() {
  if (!selectedParentNode) { showStatus('⚠️ Select a parent node in the tree first', 'err'); return; }
  const name = document.getElementById('item-name').value.trim();
  if (!name)  { showStatus('⚠️ Enter item name', 'err'); return; }

  const tplSel = document.getElementById('item-tpl-sel');
  const tplOpt = tplSel.options[tplSel.selectedIndex];
  const templateId   = tplSel.value;
  const templateName = tplOpt?.dataset?.tplname || templateId;
  if (!templateId) { showStatus('⚠️ Select a template', 'err'); return; }

  showStatus('⏳ Creating item in Sitecore…', 'info', 30000);

  try {
    const fields = {};
    if (currentTpl) currentTpl.fields.forEach(f => { if (mappings[f.id]) fields[f.id] = mappings[f.id]; });

    const parentPath = selectedParentNode.path;
    let newItemId;

    if (scConfig.apiMode === 'ssc') {
      // SSC create — parentPath as query param
      const url = buildSscUrl(null, `?path=${encodeURIComponent(parentPath)}&database=master`);
      const res = await fetch(url, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body:    JSON.stringify({ ItemName: name, TemplateID: templateId, Fields: fields })
      });
      if (!res.ok) {
        const t = await res.text();
        throw new Error(`Create failed (${res.status}): ${t.slice(0, 200)}`);
      }
      const created = await res.json();
      newItemId = created.ItemID || created.id;

    } else {
      // GraphQL mutation (XM Cloud Management API style)
      // Fall back to SSC even in graphql mode for mutations (GraphQL is read-only in many setups)
      const url = buildSscUrl(null, `?path=${encodeURIComponent(parentPath)}&database=master`);
      const res = await fetch(url, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body:    JSON.stringify({ ItemName: name, TemplateID: templateId, Fields: fields })
      });
      if (!res.ok) {
        const t = await res.text();
        throw new Error(`Create failed (${res.status}): ${t.slice(0, 200)}`);
      }
      const created = await res.json();
      newItemId = created.ItemID || created.id;
    }

    showStatus(`✅ "${name}" created successfully!`, 'ok', 5000);

    // Refresh the parent node in the tree to show the new item
    if (selectedParentNode) {
      selectedParentNode.childrenIds  = null;  // force reload
      selectedParentNode.expanded     = false;
      selectedParentNode.hasChildren  = true;
      await expandTreeNode(selectedParentNode.id);
      renderTree(document.getElementById('tree-search').value);
    }

    // Reset form
    document.getElementById('item-name').value = '';
    mappings = {};
    renderFields();

  } catch (e) {
    console.error('[SC Tool] Create error:', e);
    showStatus(`❌ ${e.message}`, 'err', 10000);
  }
}

// ══════════════════════════════════════════════════════════════
// OFFLINE: EXCEL IMPORT
// ══════════════════════════════════════════════════════════════
function importExcel(input) {
  const file = input.files[0];
  if (!file) return;
  showStatus('⏳ Reading file…', 'info', 30000);

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const tfSheet = wb.Sheets['TemplateFields'];
      if (!tfSheet) throw new Error('"TemplateFields" sheet not found');

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

      if (!parsed.length) throw new Error('No templates found in TemplateFields sheet');
      templates = parsed;
      renderTemplates();
      showStatus(`✅ ${templates.length} templates loaded!`, 'ok', 5000);
    } catch (err) {
      showStatus(`❌ ${err.message}`, 'err', 10000);
    }
  };
  reader.onerror = () => showStatus('❌ Could not read file', 'err');
  reader.readAsArrayBuffer(file);
  input.value = '';
}

// ══════════════════════════════════════════════════════════════
// OFFLINE: TEMPLATES LIST
// ══════════════════════════════════════════════════════════════
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
    div.addEventListener('click', () => selectTemplateOffline(t));
    list.appendChild(div);
  });
}

function filterTpl(val) { renderTemplates(val); }

function selectTemplateOffline(tpl) {
  currentTpl  = tpl;
  currentItem = null;
  mappings    = {};
  renderTemplates(document.getElementById('tpl-search').value);

  // Show item section in offline mode
  document.getElementById('item-section').style.display = '';
  document.getElementById('item-tpl-lbl').textContent   = tpl.name;
  document.getElementById('create-under-row').style.display = 'none';
  document.getElementById('tpl-picker-row').style.display   = 'none';
  document.getElementById('create-sc-btn').style.display    = 'none';
  document.getElementById('add-structure-btn').style.display = '';
  document.getElementById('item-name').value = '';
  setIsPage(currentItem ? currentItem.isPage : true);
  renderParentSelect();
  renderFields();
}

// ══════════════════════════════════════════════════════════════
// ITEM FORM — SHARED
// ══════════════════════════════════════════════════════════════
function setIsPage(val) {
  document.getElementById('tog-yes').className = 'tog' + (val  ? ' yes-on' : '');
  document.getElementById('tog-no').className  = 'tog' + (!val ? ' no-on'  : '');
  document.getElementById('ispage-hint').textContent = val
    ? '📄 Creates a page in Sitecore'
    : '🗂 Datasource — select which page it belongs to';
  document.getElementById('ispage-hint').className   = 'hint ' + (val ? 'blue' : 'orange');
  document.getElementById('parent-row').style.display = (!val && !isLiveMode) ? '' : 'none';
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
    nameSpan.className = 'field-name'; nameSpan.title = f.id; nameSpan.textContent = f.id;

    const typeSpan = document.createElement('span');
    typeSpan.className = 'field-type-lbl'; typeSpan.textContent = f.type;

    const inp = document.createElement('input');
    inp.className = 'field-val'; inp.type = 'text'; inp.placeholder = 'value…'; inp.value = val;
    inp.addEventListener('input', () => { mappings[f.id] = inp.value; updateFieldProgress(); });

    const btn = document.createElement('button');
    btn.className   = 'field-map-btn' + (activeFieldId === f.id ? ' active' : '');
    btn.title       = 'Pick from page';
    btn.textContent = '📌';
    btn.addEventListener('click', () => setActiveField(f.id, btn));

    div.appendChild(nameSpan); div.appendChild(typeSpan);
    div.appendChild(inp);      div.appendChild(btn);
    list.appendChild(div);
  });
}

function setActiveField(fieldId, btn) {
  if (activeFieldId === fieldId) { activeFieldId = null; btn.classList.remove('active'); return; }
  document.querySelectorAll('.field-map-btn').forEach(b => b.classList.remove('active'));
  activeFieldId = fieldId; btn.classList.add('active');
  if (!pickerActive) togglePicker();
  showStatus('⚡ Click any element on the page', 'info');
}

// ══════════════════════════════════════════════════════════════
// OFFLINE: STRUCTURE
// ══════════════════════════════════════════════════════════════
function addToStructure() {
  const name = document.getElementById('item-name').value.trim();
  if (!name || !currentTpl) { showStatus('⚠️ Enter item name and select a template', 'err'); return; }
  const isPage   = document.getElementById('tog-yes').classList.contains('yes-on');
  const parentId = document.getElementById('parent-sel').value;
  const fields   = currentTpl.fields.map(f => ({ id: f.id, type: f.type, value: mappings[f.id] || '' }));

  if (currentItem?._idx !== undefined) {
    structureList[currentItem._idx] = { ...structureList[currentItem._idx], name, template: currentTpl.name, isPage, parentId, fields };
    showStatus('✅ Item updated!', 'ok');
  } else {
    const safeId = `{NEW-${String(structureList.length + 1).padStart(2,'0')}-${name.replace(/\W+/g,'-').toUpperCase().slice(0,12)}}`;
    structureList.push({ id: safeId, name, template: currentTpl.name, isPage, parentId, fields });
    showStatus('✅ Added to structure!', 'ok');
  }

  renderStructureTree();
  document.getElementById('item-name').value = '';
  mappings = {}; currentItem = null; renderFields();
}

function saveFields() {
  if (!currentItem || currentItem._idx === undefined) { addToStructure(); return; }
  structureList[currentItem._idx].fields = currentTpl.fields.map(f => ({ id: f.id, type: f.type, value: mappings[f.id] || '' }));
  renderStructureTree();
  showStatus('✅ Fields saved!', 'ok');
}

function renderStructureTree() {
  const tree  = document.getElementById('struct-tree');
  const badge = document.getElementById('struct-count');
  badge.textContent   = structureList.length;
  badge.style.display = structureList.length ? '' : 'none';
  if (!structureList.length) { tree.innerHTML = '<div class="empty-msg">Select a template above and add items here.</div>'; return; }
  tree.innerHTML = '';
  structureList.forEach((item, i) => {
    const div = document.createElement('div');
    div.className = 'struct-item' + (currentItem?._idx === i ? ' active' : '');
    const info = document.createElement('div');
    info.style.minWidth = '0';
    info.innerHTML = `<div class="si-name">${item.isPage ? '📄' : '🗂'} ${item.name}</div><div class="si-tpl">${item.template}</div>`;
    const actions = document.createElement('div');
    actions.className = 'si-actions';
    const editBtn = document.createElement('button');
    editBtn.className = 'si-btn'; editBtn.title = 'Edit'; editBtn.textContent = '✏️';
    editBtn.addEventListener('click', e => { e.stopPropagation(); editItem(i); });
    const delBtn = document.createElement('button');
    delBtn.className = 'si-btn'; delBtn.title = 'Remove'; delBtn.textContent = '🗑️';
    delBtn.addEventListener('click', e => { e.stopPropagation(); removeItem(i); });
    actions.appendChild(editBtn); actions.appendChild(delBtn);
    div.appendChild(info); div.appendChild(actions);
    tree.appendChild(div);
  });
}

function editItem(i) {
  const item = structureList[i];
  const tpl  = templates.find(t => t.name === item.template);
  if (!tpl) { showStatus('⚠️ Template not found', 'err'); return; }
  currentTpl = tpl; currentItem = { ...item, _idx: i }; mappings = {};
  (item.fields || []).forEach(f => { mappings[f.id] = f.value || ''; });
  renderTemplates(document.getElementById('tpl-search').value);
  selectTemplateOffline(tpl);
  document.getElementById('item-name').value = item.name;
  setIsPage(item.isPage); renderParentSelect(); renderFields();
}

function removeItem(i) {
  if (!confirm(`Remove "${structureList[i].name}"?`)) return;
  structureList.splice(i, 1);
  if (currentItem?._idx === i) currentItem = null;
  renderStructureTree();
}

// ══════════════════════════════════════════════════════════════
// OFFLINE: EXPORT EXCEL
// ══════════════════════════════════════════════════════════════
function exportExcel() {
  if (!structureList.length) { showStatus('⚠️ Add at least one item first', 'err'); return; }
  try {
    const wb = XLSX.utils.book_new();
    const itemHeader = ['ID','ParentID','Name','Template','Path','IsPage','DatasourceIDs','SortOrder','Language','Version','WorkflowState','Publish'];
    const itemRows   = [itemHeader];
    structureList.forEach((item, i) => {
      itemRows.push([item.id||`{ITEM-${i+1}-ID}`,item.parentId||'',item.name||item.template,item.template,item.path||'',item.isPage?'Yes':'No',item.datasourceId||'',(i+1)*100,'en',1,'Draft','Yes']);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(itemRows), 'Items');

    const maxF = Math.max(...structureList.map(it => (it.fields||[]).length), 1);
    const fHeader = ['ItemID','Language','Version'];
    for (let i = 0; i < maxF; i++) fHeader.push('FieldName','FieldValue','FieldType');
    const fRows = [fHeader];
    structureList.forEach((item, i) => {
      const row = [item.id||`{ITEM-${i+1}-ID}`,'en',1];
      (item.fields||[]).forEach(f => row.push(f.id, f.value||'', f.type));
      while (row.length < 3 + maxF * 3) row.push('','','');
      fRows.push(row);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fRows), 'Fields');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['MediaID','FileName','FilePath','DestinationFolder','Alt','Title','Language']]), 'Media');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['ItemID','FieldName','LinkType','LinkText','Url','Target','Anchor','Class']]), 'Links');
    XLSX.writeFile(wb, 'SitecoreContentExport.xlsx');
    showStatus('✅ Exported successfully!', 'ok');
  } catch (err) {
    showStatus(`❌ Export failed: ${err.message}`, 'err');
  }
}

// ══════════════════════════════════════════════════════════════
// SECTION COLLAPSE
// ══════════════════════════════════════════════════════════════
function toggleSection(id) {
  const body  = document.getElementById(`${id}-body`);
  const arrow = document.getElementById(`${id}-arrow`);
  const open  = body.style.display !== 'none';
  body.style.display = open ? 'none' : '';
  if (arrow) arrow.textContent = open ? '▶' : '▼';
}

// ══════════════════════════════════════════════════════════════
// STATUS BAR
// ══════════════════════════════════════════════════════════════
let _st;
function showStatus(msg, type = 'info', duration = 3000) {
  const bar = document.getElementById('status-bar');
  bar.textContent = msg; bar.className = `status-bar ${type}`; bar.style.display = 'block';
  clearTimeout(_st);
  if (duration > 0) _st = setTimeout(() => { bar.style.display = 'none'; }, duration);
}
