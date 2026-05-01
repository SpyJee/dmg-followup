const https = require('https');
const crypto = require('crypto');

// ══════════════════════════════════════════
// DMG Aerospace — Production Follow-Up API
// Routes:
//   POST /api/followup/{order-save}      — Save/update order
//   GET  /api/followup/{order-list}      — List all orders
//   POST /api/followup/{order-delete}    — Delete order
//   POST /api/followup/{recipe-save}     — Save/update recipe
//   GET  /api/followup/{recipe-list}     — List all recipes
//   POST /api/followup/{recipe-delete}   — Delete recipe
//   POST /api/followup/{template-save}   — Save/update operation template
//   GET  /api/followup/{template-list}   — List all templates
//   POST /api/followup/{template-delete} — Delete template
//   POST /api/followup/{method-todo-update} — Toggle one method-method-todo box
//   GET  /api/followup/{method-todo-list}   — List all method-todo entries
//   POST /api/followup/{suivi-update}    — Update operation status
//   GET  /api/followup/{suivi-list}      — Get all suivi statuses
//   GET  /api/followup/{holidays}        — Get holiday list
//   GET  /api/followup/{health}          — Health check
//   GET  /api/followup/{po-pdf}          — Redirect to a SAS URL for the PO PDF in dmgpdfvault
// ══════════════════════════════════════════

// Tables:
//   followuporders     — PartitionKey: 'ORDER', RowKey: commandeItem
//   followuprecipes    — PartitionKey: 'RECIPE', RowKey: partNumber
//   followupsuivi      — PartitionKey: 'SUIVI', RowKey: commandeItem
//   followuptemplates  — PartitionKey: 'TEMPLATE', RowKey: name
//   followupmethodtodos — PartitionKey: 'METHOD', RowKey: commandeItem

module.exports = async function (context, req) {
  const cors = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Content-Type': 'application/json'
  };

  if (req.method === 'OPTIONS') {
    context.res = { status: 204, headers: cors, body: '' };
    return;
  }

  const route = req.params.route || '';
  context.log(`Follow-Up Function: ${req.method} /${route}`);

  try {
    if (route === 'health')        { context.res = { status: 200, headers: cors, body: JSON.stringify({ status: 'ok', node: process.version, timestamp: new Date().toISOString() }) }; return; }

    // Auth gate — enforced only when AUTH_REQUIRED=1 in app settings.
    if (process.env.AUTH_REQUIRED === '1' && route !== 'health') {
      const session = await authReadSession(parseBearer(req));
      if (!session) {
        context.res = { status: 401, headers: cors, body: JSON.stringify({ error: 'Authentication required', code: 'auth_required' }) };
        return;
      }
      req.authUser = session;
    }

    if (route === 'order-save')    { await handleOrderSave(context, req, cors); return; }
    if (route === 'order-list')    { await handleOrderList(context, req, cors); return; }
    if (route === 'order-delete')  { await handleOrderDelete(context, req, cors); return; }
    if (route === 'recipe-save')   { await handleRecipeSave(context, req, cors); return; }
    if (route === 'recipe-list')   { await handleRecipeList(context, req, cors); return; }
    if (route === 'recipe-delete') { await handleRecipeDelete(context, req, cors); return; }
    if (route === 'template-save')   { await handleTemplateSave(context, req, cors); return; }
    if (route === 'template-list')   { await handleTemplateList(context, req, cors); return; }
    if (route === 'template-delete') { await handleTemplateDelete(context, req, cors); return; }
    if (route === 'method-todo-update') { await handleMethodTodoUpdate(context, req, cors); return; }
    if (route === 'method-todo-list')   { await handleMethodTodoList(context, req, cors); return; }
    if (route === 'suivi-update')  { await handleSuiviUpdate(context, req, cors); return; }
    if (route === 'suivi-list')    { await handleSuiviList(context, req, cors); return; }
    if (route === 'holidays')      { handleHolidays(context, cors); return; }
    if (route === 'po-pdf')        { await handlePoPdf(context, req, cors); return; }

    context.res = { status: 404, headers: cors, body: JSON.stringify({ error: 'Unknown route: ' + route }) };
  } catch (err) {
    context.log.error('Handler error:', err.message, err.stack);
    context.res = { status: 500, headers: cors, body: JSON.stringify({ error: err.message }) };
  }
};

// ══════════════════════════════════════════
// ORDER ROUTES
// ══════════════════════════════════════════

async function handleOrderSave(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followuporders');

  const o = req.body;
  if (!o || !o.commandeItem) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'commandeItem required' }) };
    return;
  }

  const entity = {
    PartitionKey: 'ORDER',
    RowKey: o.commandeItem,
    po: o.po || '',
    item: o.item || '',
    commandeItem: o.commandeItem,
    client: o.client || '',
    partNumber: o.partNumber || '',
    dateRequise: o.dateRequise || '',
    qty: o.qty || 1,
    workOrder: o.workOrder || '',
    comment: o.comment || '',
    buyer: o.buyer || '',
    archived: o.archived ? 'true' : 'false',
    carcoStatus: o.carcoStatus || '',
    carcoEcd: o.carcoEcd || '',
    carcoNote: o.carcoNote || '',
    createdAt: o.createdAt || new Date().toISOString()
  };

  await tableUpsert(account, key, 'followuporders', entity);
  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true, commandeItem: o.commandeItem }) };
}

async function handleOrderList(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followuporders');

  const filter = encodeURIComponent("PartitionKey eq 'ORDER'");
  const raw = await tableReq(account, key, 'GET', `/followuporders()?$filter=${filter}&$top=1000`);
  const data = JSON.parse(raw);
  const orders = (data.value || []).map(e => ({
    po: e.po, item: e.item, commandeItem: e.RowKey,
    client: e.client, partNumber: e.partNumber,
    dateRequise: e.dateRequise, qty: e.qty,
    workOrder: e.workOrder, comment: e.comment,
    buyer: e.buyer || '',
    archived: e.archived === 'true',
    carcoStatus: e.carcoStatus || '', carcoEcd: e.carcoEcd || '', carcoNote: e.carcoNote || '',
    createdAt: e.createdAt
  }));

  context.res = { status: 200, headers: cors, body: JSON.stringify(orders) };
}

async function handleOrderDelete(context, req, cors) {
  const { account, key } = getStorage();
  const commandeItem = (req.body && req.body.commandeItem) || '';
  if (!commandeItem) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'commandeItem required' }) };
    return;
  }

  try {
    await tableReq(account, key, 'DELETE', `/followuporders(PartitionKey='ORDER',RowKey='${encodeURIComponent(commandeItem)}')`, null, { 'If-Match': '*' });
  } catch (e) { /* ignore 404 */ }

  // Also delete suivi record
  try {
    await tableReq(account, key, 'DELETE', `/followupsuivi(PartitionKey='SUIVI',RowKey='${encodeURIComponent(commandeItem)}')`, null, { 'If-Match': '*' });
  } catch (e) { /* ignore */ }

  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true, deleted: commandeItem }) };
}

// ══════════════════════════════════════════
// RECIPE ROUTES
// ══════════════════════════════════════════

async function handleRecipeSave(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followuprecipes');

  const r = req.body;
  if (!r || !r.partNumber) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'partNumber required' }) };
    return;
  }

  const entity = {
    PartitionKey: 'RECIPE',
    RowKey: r.partNumber,
    partNumber: r.partNumber,
    template: r.template || 'Custom',
    opsJson: JSON.stringify(r.ops || []),
    nbOps: r.nbOps || (r.ops || []).length,
    totalDelay: r.totalDelay || (r.ops || []).reduce((s, o) => s + (o.delay || 0), 0)
  };

  await tableUpsert(account, key, 'followuprecipes', entity);
  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true, partNumber: r.partNumber }) };
}

async function handleRecipeList(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followuprecipes');

  const filter = encodeURIComponent("PartitionKey eq 'RECIPE'");
  const raw = await tableReq(account, key, 'GET', `/followuprecipes()?$filter=${filter}&$top=1000`);
  const data = JSON.parse(raw);
  const recipes = (data.value || []).map(e => ({
    partNumber: e.RowKey,
    template: e.template,
    ops: JSON.parse(e.opsJson || '[]'),
    nbOps: e.nbOps,
    totalDelay: e.totalDelay
  }));

  context.res = { status: 200, headers: cors, body: JSON.stringify(recipes) };
}

async function handleRecipeDelete(context, req, cors) {
  const { account, key } = getStorage();
  const partNumber = (req.body && req.body.partNumber) || '';
  if (!partNumber) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'partNumber required' }) };
    return;
  }

  try {
    await tableReq(account, key, 'DELETE', `/followuprecipes(PartitionKey='RECIPE',RowKey='${encodeURIComponent(partNumber)}')`, null, { 'If-Match': '*' });
  } catch (e) { /* ignore 404 */ }

  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true, deleted: partNumber }) };
}

// ══════════════════════════════════════════
// TEMPLATE ROUTES
// ══════════════════════════════════════════

async function handleTemplateSave(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followuptemplates');

  const t = req.body;
  if (!t || !t.name) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'name required' }) };
    return;
  }

  const entity = {
    PartitionKey: 'TEMPLATE',
    RowKey: t.name,
    name: t.name,
    opsJson: JSON.stringify(t.ops || [])
  };

  await tableUpsert(account, key, 'followuptemplates', entity);
  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true, name: t.name }) };
}

async function handleTemplateList(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followuptemplates');

  const filter = encodeURIComponent("PartitionKey eq 'TEMPLATE'");
  const raw = await tableReq(account, key, 'GET', `/followuptemplates()?$filter=${filter}&$top=1000`);
  const data = JSON.parse(raw);
  const templates = (data.value || []).map(e => ({
    name: e.RowKey,
    ops: JSON.parse(e.opsJson || '[]')
  }));

  context.res = { status: 200, headers: cors, body: JSON.stringify(templates) };
}

async function handleTemplateDelete(context, req, cors) {
  const { account, key } = getStorage();
  const name = (req.body && req.body.name) || '';
  if (!name) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'name required' }) };
    return;
  }

  try {
    await tableReq(account, key, 'DELETE', `/followuptemplates(PartitionKey='TEMPLATE',RowKey='${encodeURIComponent(name)}')`, null, { 'If-Match': '*' });
  } catch (e) { /* ignore 404 */ }

  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true, deleted: name }) };
}

// ══════════════════════════════════════════
// METHOD TO-DO ROUTES
// ══════════════════════════════════════════
// 5 fixed boxes per order: recipe, programming, drawing, dimReg, range.
// Each box stores { by, at } when checked. Order auto-clears from list
// once all 5 are populated (handled in the frontend filter).

async function handleMethodTodoUpdate(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followupmethodtodos');

  const t = req.body || {};
  const VALID = ['recipe', 'programming', 'drawing', 'dimReg', 'range'];
  if (!t.commandeItem || !VALID.includes(t.box)) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'commandeItem + valid box required' }) };
    return;
  }

  const username = (req.authUser && req.authUser.username) || t.by || 'unknown';
  const path = `/followupmethodtodos(PartitionKey='METHOD',RowKey='${encodeURIComponent(t.commandeItem)}')`;
  let existing = {};
  try {
    const raw = await tableReq(account, key, 'GET', path);
    existing = JSON.parse(raw);
  } catch (_) { /* row doesn't exist yet */ }

  const boxes = JSON.parse(existing.boxesJson || '{}');
  if (t.checked === false) {
    delete boxes[t.box];
  } else {
    boxes[t.box] = { by: username, at: new Date().toISOString() };
  }

  const entity = {
    PartitionKey: 'METHOD',
    RowKey: t.commandeItem,
    commandeItem: t.commandeItem,
    boxesJson: JSON.stringify(boxes)
  };
  await tableUpsert(account, key, 'followupmethodtodos', entity);
  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true, boxes }) };
}

async function handleMethodTodoList(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followupmethodtodos');

  const filter = encodeURIComponent("PartitionKey eq 'METHOD'");
  // Azure Table Storage caps $top at 1000 — anything higher is a 400.
  const raw = await tableReq(account, key, 'GET', `/followupmethodtodos()?$filter=${filter}&$top=1000`);
  const data = JSON.parse(raw);
  const todos = (data.value || []).map(e => ({
    commandeItem: e.RowKey,
    boxes: JSON.parse(e.boxesJson || '{}')
  }));

  context.res = { status: 200, headers: cors, body: JSON.stringify(todos) };
}

// ══════════════════════════════════════════
// SUIVI ROUTES
// ══════════════════════════════════════════

async function handleSuiviUpdate(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followupsuivi');

  const s = req.body;
  if (!s || !s.commandeItem) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'commandeItem required' }) };
    return;
  }

  // s.statuses = { op1: 'Requis', op2: 'En cours', ... }
  const entity = {
    PartitionKey: 'SUIVI',
    RowKey: s.commandeItem,
    commandeItem: s.commandeItem,
    statusesJson: JSON.stringify(s.statuses || {})
  };

  await tableUpsert(account, key, 'followupsuivi', entity);
  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true, commandeItem: s.commandeItem }) };
}

async function handleSuiviList(context, req, cors) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'followupsuivi');

  const filter = encodeURIComponent("PartitionKey eq 'SUIVI'");
  const raw = await tableReq(account, key, 'GET', `/followupsuivi()?$filter=${filter}&$top=1000`);
  const data = JSON.parse(raw);
  const result = {};
  (data.value || []).forEach(e => {
    result[e.RowKey] = JSON.parse(e.statusesJson || '{}');
  });

  context.res = { status: 200, headers: cors, body: JSON.stringify(result) };
}

// ══════════════════════════════════════════
// HOLIDAYS
// ══════════════════════════════════════════

function handleHolidays(context, cors) {
  const holidays = [
    // 2026
    '2026-01-01','2026-04-03','2026-05-18','2026-06-24','2026-07-01',
    '2026-07-20','2026-07-21','2026-07-22','2026-07-23','2026-07-24',
    '2026-07-27','2026-07-28','2026-07-29','2026-07-30','2026-07-31',
    '2026-09-07','2026-10-12','2026-12-25',
    // 2027
    '2027-01-01','2027-03-26','2027-05-24','2027-06-24','2027-07-01',
    '2027-07-19','2027-07-20','2027-07-21','2027-07-22','2027-07-23',
    '2027-07-26','2027-07-27','2027-07-28','2027-07-29','2027-07-30',
    '2027-09-06','2027-10-11','2027-12-25'
  ];
  context.res = { status: 200, headers: cors, body: JSON.stringify(holidays) };
}

// ══════════════════════════════════════════
// TABLE STORAGE HELPERS
// ══════════════════════════════════════════

function getStorage() {
  const account = process.env.STORAGE_ACCOUNT_NAME;
  const key = process.env.STORAGE_ACCOUNT_KEY;
  if (!account || !key) throw new Error('STORAGE_ACCOUNT_NAME or STORAGE_ACCOUNT_KEY not configured');
  return { account, key };
}

function callHttps(host, path, method, headers, body) {
  return new Promise((resolve, reject) => {
    const opts = { hostname: host, path, method, headers };
    const r = https.request(opts, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) reject(new Error(`${res.statusCode}: ${data}`));
        else resolve(data);
      });
    });
    r.on('error', reject);
    if (body) r.write(body);
    r.end();
  });
}

async function tableReq(account, key, method, path, body, extraHeaders) {
  const date = new Date().toUTCString();
  const ct = body ? 'application/json' : '';
  const cl = body ? Buffer.byteLength(JSON.stringify(body)).toString() : '0';

  const pathOnly = path.split('?')[0];
  const stringToSign = [method, '', ct, date, `/${account}${pathOnly}`].join('\n');
  const sig = crypto.createHmac('sha256', Buffer.from(key, 'base64'))
    .update(stringToSign, 'utf8').digest('base64');

  const headers = {
    'Authorization': `SharedKey ${account}:${sig}`,
    'x-ms-date': date,
    'x-ms-version': '2019-02-02',
    'Accept': 'application/json;odata=nometadata',
    'DataServiceVersion': '3.0;NetFx',
  };
  if (body) {
    headers['Content-Type'] = 'application/json';
    headers['Content-Length'] = cl;
  }
  if (extraHeaders) Object.assign(headers, extraHeaders);

  return callHttps(`${account}.table.core.windows.net`, path, method, headers,
    body ? JSON.stringify(body) : null);
}

async function ensureTable(account, key, name) {
  try { await tableReq(account, key, 'POST', '/Tables', { TableName: name }); }
  catch (e) { /* already exists */ }
}

async function tableUpsert(account, key, table, entity) {
  // Try insert first, then merge on conflict
  try {
    await tableReq(account, key, 'POST', `/${table}`, entity);
  } catch (e) {
    if (e.message && e.message.includes('409')) {
      // Entity exists — merge update
      const pk = encodeURIComponent(entity.PartitionKey);
      const rk = encodeURIComponent(entity.RowKey);
      await tableReq(account, key, 'PUT', `/${table}(PartitionKey='${pk}',RowKey='${rk}')`, entity, { 'If-Match': '*' });
    } else {
      throw e;
    }
  }
}


// ══════════════════════════════════════════
// PO PDF lookup — redirect to a SAS URL pointing at dmgpdfvault/po-pdfs.
// The on-prem sync (sync-pos-to-blob.ps1) keeps the container in step with
// T:\Bon de Commande every hour; this endpoint just searches the cached
// blob list for paths containing the PO and hands the user a short SAS.
// ══════════════════════════════════════════

const PO_BLOB_CONTAINER = process.env.PO_BLOB_CONTAINER || 'po-pdfs';
const PO_USER_SAS_TTL_MIN = parseInt(process.env.PO_USER_SAS_TTL_MIN || '10', 10);

// In-memory cache for the blob list. Resets on cold start; refreshed every 5 min.
let _poBlobCache = { ts: 0, paths: [] };
const PO_CACHE_TTL_MS = 5 * 60 * 1000;

async function handlePoPdf(context, req, cors) {
  const po = String((req.query && req.query.po) || '').trim();
  const site = String((req.query && req.query.site) || 'DMG').toUpperCase();
  const username = (req.authUser && req.authUser.username) || 'unauth';
  const ip = clientIp(req);

  if (!po) {
    await poAuditLog(context, { action: 'lookup', user: username, po: '', site, ip, success: false, message: 'missing po' });
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'po required' }) };
    return;
  }

  const paths = await getPoBlobPaths(context);
  const match = pickBestPoBlob(paths, po, site);
  if (!match) {
    await poAuditLog(context, { action: 'lookup', user: username, po, site, ip, success: false, message: 'not found' });
    context.res = { status: 404, headers: cors, body: JSON.stringify({ error: 'PO not found in cloud cache. The hourly sync may not have run yet.' }) };
    return;
  }

  const { account, key } = getPoBlobsStorage();
  const sasUrl = generateBlobSas(account, key, PO_BLOB_CONTAINER, match, 'r', PO_USER_SAS_TTL_MIN * 60);
  await poAuditLog(context, { action: 'lookup', user: username, po, site, ip, success: true, message: match });
  // Return JSON, not 302: anchor navigation strips the bearer header, so the
  // frontend has to fetch with bearer then window.open the SAS itself.
  context.res = { status: 200, headers: cors, body: JSON.stringify({ sasUrl, blobPath: match, expiresInSec: PO_USER_SAS_TTL_MIN * 60 }) };
}

// Pick the most-likely PDF when several blobs match the same PO. Mirrors the
// dmgpo:// launcher heuristic: prefer "PO <num>*.pdf", then *<po>*.pdf with
// shortest filename, then any path containing the PO. Stable & deterministic.
function pickBestPoBlob(paths, po, site) {
  const lowerPo = po.toLowerCase();
  const sitePrefix = site.toUpperCase() + '/';
  const candidates = paths.filter(p =>
    p.toUpperCase().startsWith(sitePrefix) && p.toLowerCase().includes(lowerPo));
  if (!candidates.length) return null;

  function fileBase(p) {
    const f = p.split('/').pop() || '';
    return f.toLowerCase().replace(/\.pdf$/i, '');
  }
  // Tier 1: filename starts with "PO " + po (most explicit)
  const tier1 = candidates.filter(p => fileBase(p).startsWith('po ' + lowerPo));
  if (tier1.length) return tier1.sort((a, b) => a.length - b.length)[0];
  // Tier 2: filename equals po (exact match)
  const tier2 = candidates.filter(p => fileBase(p) === lowerPo);
  if (tier2.length) return tier2[0];
  // Tier 3: filename starts with po
  const tier3 = candidates.filter(p => fileBase(p).startsWith(lowerPo));
  if (tier3.length) return tier3.sort((a, b) => a.length - b.length)[0];
  // Tier 4: any path containing po — pick shortest
  return candidates.sort((a, b) => a.length - b.length)[0];
}

async function getPoBlobPaths(context) {
  const now = Date.now();
  if (_poBlobCache.paths.length && (now - _poBlobCache.ts) < PO_CACHE_TTL_MS) {
    return _poBlobCache.paths;
  }
  const { account, key } = getPoBlobsStorage();
  const all = [];
  let marker = null;
  while (true) {
    const qs = new URLSearchParams({ restype: 'container', comp: 'list', maxresults: '5000' });
    if (marker) qs.set('marker', marker);
    const xml = await blobReq(account, key, 'GET', `/${PO_BLOB_CONTAINER}?${qs.toString()}`);
    const blobBlocks = xml.match(/<Blob>[\s\S]*?<\/Blob>/g) || [];
    for (const b of blobBlocks) {
      const name = (b.match(/<Name>([\s\S]*?)<\/Name>/) || [])[1];
      if (name) all.push(name);
    }
    const next = (xml.match(/<NextMarker>([\s\S]*?)<\/NextMarker>/) || [])[1];
    if (!next || !next.trim()) break;
    marker = next.trim();
  }
  _poBlobCache = { ts: now, paths: all };
  context.log(`po-pdf: refreshed blob cache, ${all.length} paths`);
  return all;
}

// ----- PO PDF helpers -------------------------------------------------------

function getPoBlobsStorage() {
  const account = process.env.POBLOBS_ACCOUNT_NAME;
  const key = process.env.POBLOBS_ACCOUNT_KEY;
  if (!account || !key) throw new Error('POBLOBS_ACCOUNT_NAME/KEY not configured');
  return { account, key };
}

function clientIp(req) {
  const h = req.headers || {};
  const xff = h['x-forwarded-for'] || h['X-Forwarded-For'] || '';
  if (xff) return String(xff).split(',')[0].trim().split(':')[0];
  return h['x-azure-clientip'] || '';
}

// Service SAS for one specific blob. Permissions: 'r' (read) or 'cw' (create+write).
function generateBlobSas(account, key, container, blobName, permissions, ttlSec) {
  const sv = '2020-08-04';
  const sr = 'b';
  const sp = permissions;
  const se = new Date(Date.now() + ttlSec * 1000).toISOString().replace(/\.\d+Z$/, 'Z');
  const spr = 'https';
  const canonicalizedResource = `/blob/${account}/${container}/${blobName}`;
  const stringToSign = [
    sp, '', se, canonicalizedResource,
    '', '', spr, sv, sr,
    '', '', '', '', '', '', ''
  ].join('\n');
  const sig = crypto.createHmac('sha256', Buffer.from(key, 'base64'))
    .update(stringToSign, 'utf8').digest('base64');
  const params = new URLSearchParams({ sv, sr, sp, se, spr, sig });
  const encodedBlob = blobName.split('/').map(encodeURIComponent).join('/');
  return `https://${account}.blob.core.windows.net/${container}/${encodedBlob}?${params.toString()}`;
}

// Blob SharedKey-signed REST request. Returns response body string (XML for list).
function blobReq(account, key, method, pathWithQuery) {
  return new Promise((resolve, reject) => {
    const date = new Date().toUTCString();
    const [pathOnly, qs = ''] = pathWithQuery.split('?');
    const ch = `x-ms-date:${date}\nx-ms-version:2020-08-04\n`;
    const cr = `/${account}${pathOnly}` + canonicalizeBlobQuery(qs);
    const stringToSign = [method, '', '', '', '', '', '', '', '', '', '', '', ch + cr].join('\n');
    const sig = crypto.createHmac('sha256', Buffer.from(key, 'base64'))
      .update(stringToSign, 'utf8').digest('base64');
    const opts = {
      hostname: `${account}.blob.core.windows.net`, path: pathWithQuery, method,
      headers: {
        'Authorization': `SharedKey ${account}:${sig}`,
        'x-ms-date': date,
        'x-ms-version': '2020-08-04',
        'Content-Length': '0'
      }
    };
    const r = https.request(opts, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) reject(new Error(`${res.statusCode}: ${data.slice(0, 300)}`));
        else resolve(data);
      });
    });
    r.on('error', reject);
    r.end();
  });
}

function canonicalizeBlobQuery(qs) {
  if (!qs) return '';
  const params = new URLSearchParams(qs);
  const map = {};
  for (const [k, v] of params.entries()) {
    const lk = k.toLowerCase();
    map[lk] = (map[lk] ? map[lk] + ',' : '') + v;
  }
  return Object.keys(map).sort().map(k => `\n${k}:${map[k]}`).join('');
}

async function poAuditLog(context, evt) {
  try {
    const { account, key } = getStorage();
    await ensureTable(account, key, 'poaudit');
    const now = new Date();
    const day = now.toISOString().slice(0, 10);
    const rk = `${now.getTime().toString().padStart(15, '0')}-${crypto.randomUUID().slice(0, 8)}`;
    const entity = {
      PartitionKey: day,
      RowKey: rk,
      action: evt.action || '',
      user: evt.user || '',
      po: evt.po || '',
      site: evt.site || '',
      ip: evt.ip || '',
      success: !!evt.success,
      message: (evt.message || '').slice(0, 500),
      at: now.toISOString()
    };
    await tableReq(account, key, 'POST', '/poaudit', entity);
  } catch (e) {
    context.log.warn('poAuditLog failed:', e.message);
  }
}

// ══════════════════════════════════════════
// AUTH — reads dmgsessions from the shared auth storage account
// (set AUTH_STORAGE_ACCOUNT_NAME / AUTH_STORAGE_ACCOUNT_KEY on this
// Function App's config; falls back to STORAGE_* if not set, which
// only works when the follow-up and auth storage happen to match).
// ══════════════════════════════════════════

function getAuthStorage() {
  const account = process.env.AUTH_STORAGE_ACCOUNT_NAME || process.env.STORAGE_ACCOUNT_NAME;
  const key = process.env.AUTH_STORAGE_ACCOUNT_KEY || process.env.STORAGE_ACCOUNT_KEY;
  if (!account || !key) throw new Error('AUTH_STORAGE_ACCOUNT_NAME/KEY not configured');
  return { account, key };
}

function parseBearer(req) {
  const h = (req.headers && (req.headers.authorization || req.headers.Authorization)) || '';
  const m = /^Bearer\s+(.+)$/i.exec(h);
  return m ? m[1].trim() : null;
}

function authTokenHash(token) {
  return crypto.createHash('sha256').update(token, 'utf8').digest('hex');
}

async function authReadSession(token) {
  if (!token || typeof token !== 'string') return null;
  const { account, key } = getAuthStorage();
  const tokenHash = authTokenHash(token);
  try {
    const body = await tableReq(account, key, 'GET', `/dmgsessions(PartitionKey='session',RowKey='${tokenHash}')`);
    const s = JSON.parse(body);
    if (!s || !s.expiresAt) return null;
    if (Date.now() > Date.parse(s.expiresAt)) return null;
    return s;
  } catch (_) {
    return null;
  }
}
