const express = require('express');
const axios   = require('axios');
const cors    = require('cors');
const XLSX    = require('xlsx');
const fs      = require('fs');
const multer  = require('multer');
const os      = require('os');
const path    = require('path');

// Config: env vars (Vercel) or local config.js
let config;
try {
  config = require('./config');
} catch {
  config = {
    sitecoreUrl: process.env.SITECORE_URL      || '',
    username:    process.env.SITECORE_USERNAME  || '',
    password:    process.env.SITECORE_PASSWORD  || '',
    apiKey:      process.env.SITECORE_API_KEY   || '',
    platform:    process.env.SITECORE_PLATFORM  || 'JSS',
  };
}

const uploadDir = process.env.VERCEL ? os.tmpdir() : 'uploads';
const upload    = multer({ dest: uploadDir });

const app = express();
// Open CORS so the Chrome extension (chrome-extension://) can call this API
app.use(cors({ origin: '*' }));
app.use(express.json());

// ── Health check ─────────────────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({
    status: '✅ SC Tool Extension Backend is running',
    endpoints: [
      'POST /api/import-excel',
      'POST /api/export-excel',
      'PATCH /api/update-item',
      'POST /api/create-item'
    ]
  });
});

// ── Helper ────────────────────────────────────────────────────────────────
function withApiKey(url) {
  if (!config.apiKey) return url;
  return url + (url.includes('?') ? '&' : '?') + 'sc_apikey=' + config.apiKey;
}

// ── Import Excel ──────────────────────────────────────────────────────────
app.post('/api/import-excel', upload.single('file'), (req, res) => {
  try {
    const wb = XLSX.readFile(req.file.path);

    // TemplateFields sheet → templates
    const tfSheet = wb.Sheets['TemplateFields'];
    if (!tfSheet) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: '"TemplateFields" sheet not found' });
    }
    const tfRows    = XLSX.utils.sheet_to_json(tfSheet, { header: 1 });
    const templates = [];
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
      templates.push({ name: tplName, fields, isPage: true, datasourceIds: [], pageIds: [] });
    });

    // Items sheet → existing items
    const itemsSheet   = wb.Sheets['Items'];
    const existingItems = [];
    if (itemsSheet) {
      const itemRows = XLSX.utils.sheet_to_json(itemsSheet, { header: 1 });
      itemRows.slice(1).forEach(row => {
        const id       = String(row[0] || '').trim();
        const parentId = String(row[1] || '').trim();
        const name     = String(row[2] || '').trim();
        const template = String(row[3] || '').trim();
        const isPage   = String(row[5] || '').trim().toLowerCase() === 'yes';
        if (!template) return;
        existingItems.push({ id, parentId, name, template, isPage });
      });
    }

    fs.unlinkSync(req.file.path);
    res.json({ templates, existingItems });
  } catch (err) {
    if (req.file?.path) try { fs.unlinkSync(req.file.path); } catch {}
    res.status(500).json({ error: err.message });
  }
});

// ── Export Excel ──────────────────────────────────────────────────────────
app.post('/api/export-excel', (req, res) => {
  try {
    const { items } = req.body;
    const wb = XLSX.utils.book_new();

    // Items sheet
    const itemHeader = ['ID','ParentID','Name','Template','Path','IsPage','DatasourceIDs','SortOrder','Language','Version','WorkflowState','Publish'];
    const itemRows   = [itemHeader];
    items.forEach((item, i) => {
      itemRows.push([
        item.id        || `{ITEM-${i+1}-ID}`,
        item.parentId  || '',
        item.name      || item.template,
        item.template,
        item.path      || '',
        item.isPage    ? 'Yes' : 'No',
        item.datasourceId || '',
        (i + 1) * 100,
        'en', 1, 'Draft', 'Yes'
      ]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(itemRows), 'Items');

    // Fields sheet
    const maxFields  = Math.max(...items.map(it => (it.fields || []).length), 1);
    const fieldHeader = ['ItemID', 'Language', 'Version'];
    for (let i = 0; i < maxFields; i++) fieldHeader.push('FieldName', 'FieldValue', 'FieldType');
    const fieldRows = [fieldHeader];
    items.forEach((item, i) => {
      const row = [item.id || `{ITEM-${i+1}-ID}`, 'en', 1];
      (item.fields || []).forEach(f => row.push(f.id, f.value || '', f.type));
      while (row.length < 3 + maxFields * 3) row.push('', '', '');
      fieldRows.push(row);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fieldRows), 'Fields');

    // Media + Links templates
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['MediaID','FileName','FilePath','DestinationFolder','Alt','Title','Language']
    ]), 'Media');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['ItemID','FieldName','LinkType','LinkText','Url','Target','Anchor','Class']
    ]), 'Links');

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename="SitecoreContentExport.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Update Sitecore Item ──────────────────────────────────────────────────
app.patch('/api/update-item', async (req, res) => {
  const { id, fields } = req.body;
  if (!id) return res.status(400).json({ error: 'id required' });
  try {
    const url = withApiKey(`${config.sitecoreUrl}/sitecore/api/ssc/item/${id}`);
    await axios.patch(url, { Fields: fields }, { headers: { 'Content-Type': 'application/json' } });
    res.json({ success: true });
  } catch (err) {
    res.status(err.response?.status || 500).json({ error: err.message });
  }
});

// ── Create Sitecore Item ──────────────────────────────────────────────────
app.post('/api/create-item', async (req, res) => {
  const { name, templateId, parentPath, fields } = req.body;
  try {
    const url = withApiKey(
      `${config.sitecoreUrl}/sitecore/api/ssc/item/?path=${encodeURIComponent(parentPath || '/sitecore/content')}`
    );
    const r = await axios.post(url,
      { ItemName: name, TemplateID: templateId, Fields: fields || {} },
      { headers: { 'Content-Type': 'application/json' } }
    );
    res.json({ success: true, item: r.data });
  } catch (err) {
    res.status(err.response?.status || 500).json({ error: err.message });
  }
});

// ── Start ─────────────────────────────────────────────────────────────────
module.exports = app;

if (!process.env.VERCEL) {
  if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');
  app.listen(3002, () => {
    console.log('\n  ⚡ SC Extension Backend');
    console.log('  → http://localhost:3002\n');
  });
}
