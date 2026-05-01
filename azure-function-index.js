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
//   POST /api/followup/{po-pdf-request}     — User asks for a PO PDF
//   GET  /api/followup/{po-pdf-status}      — User polls for PO PDF readiness
//   GET  /api/followup/{po-agent-poll}      — On-prem agent polls for work
//   GET  /api/followup/{po-agent-write-sas} — Agent gets a per-blob write SAS
//   POST /api/followup/{po-agent-complete}  — Agent reports request done/failed
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
    if (route === 'po-pdf-request')     { await handlePoPdfRequest(context, req, cors); return; }
    if (route === 'po-pdf-status')      { await handlePoPdfStatus(context, req, cors); return; }
    if (route === 'po-agent-poll')      { await handlePoAgentPoll(context, req, cors); return; }
    if (route === 'po-agent-write-sas') { await handlePoAgentWriteSas(context, req, cors); return; }
    if (route === 'po-agent-complete')  { await handlePoAgentComplete(context, req, cors); return; }

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
// PO PDF ROUTES — on-demand fetch from on-prem T:\Bon de Commande
// User clicks ↗ → frontend POSTs request → on-prem agent polls,
// uploads to dmgpdfvault, proxy returns 10-min read SAS to frontend.
// Designed against PIPEDA + Loi 25 (audit log + Canadian residency).
// ══════════════════════════════════════════

const PO_VALID_SITES = ['DMG', 'DICI'];
const PO_BLOB_CONTAINER = process.env.PO_BLOB_CONTAINER || 'po-pdfs';
const PO_RATE_LIMIT_PER_HOUR = parseInt(process.env.PO_RATE_LIMIT_PER_HOUR || '100', 10);
const PO_AGENT_CLAIM_TIMEOUT_SEC = parseInt(process.env.PO_AGENT_CLAIM_TIMEOUT_SEC || '60', 10);
const PO_USER_SAS_TTL_MIN = parseInt(process.env.PO_USER_SAS_TTL_MIN || '10', 10);
const PO_AGENT_SAS_TTL_SEC = parseInt(process.env.PO_AGENT_SAS_TTL_SEC || '60', 10);

// ----- Frontend: create a request --------------------------------------------

async function handlePoPdfRequest(context, req, cors) {
  if (req.method !== 'POST') {
    context.res = { status: 405, headers: cors, body: JSON.stringify({ error: 'POST required' }) };
    return;
  }

  const body = req.body || {};
  const po = String(body.po || '').trim();
  const site = String(body.site || 'DMG').toUpperCase();
  const username = (req.authUser && req.authUser.username) || 'unauth';
  const ip = clientIp(req);

  if (!po) {
    await poAuditLog(context, { action: 'request', user: username, po: '', site, ip, success: false, message: 'missing po' });
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'po required' }) };
    return;
  }
  if (!PO_VALID_SITES.includes(site)) {
    await poAuditLog(context, { action: 'request', user: username, po, site, ip, success: false, message: 'bad site' });
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'site must be one of ' + PO_VALID_SITES.join(',') }) };
    return;
  }

  // Rate limit per user, per hour. On limit hit, audit + 429.
  const overLimit = await poRateLimitCheck(context, username);
  if (overLimit) {
    await poAuditLog(context, { action: 'request', user: username, po, site, ip, success: false, message: 'rate-limited' });
    context.res = { status: 429, headers: cors, body: JSON.stringify({ error: 'rate limit exceeded — try again later' }) };
    return;
  }

  const id = crypto.randomUUID();
  const nowIso = new Date().toISOString();
  const { account, key } = getStorage();
  await ensureTable(account, key, 'popdfrequests');

  const entity = {
    PartitionKey: site,
    RowKey: id,
    po,
    status: 'pending',
    requestedBy: username,
    requestedFromIp: ip,
    createdAt: nowIso,
    updatedAt: nowIso,
    claimedAt: '',
    blobPath: '',
    error: ''
  };
  await tableReq(account, key, 'POST', '/popdfrequests', entity);

  await poAuditLog(context, { action: 'request', user: username, po, site, ip, success: true, requestId: id });
  context.res = { status: 200, headers: cors, body: JSON.stringify({ id, status: 'pending' }) };
}

// ----- Frontend: poll status --------------------------------------------------

async function handlePoPdfStatus(context, req, cors) {
  const id = (req.query && req.query.id) || '';
  const site = String((req.query && req.query.site) || 'DMG').toUpperCase();
  const username = (req.authUser && req.authUser.username) || 'unauth';

  if (!id || !PO_VALID_SITES.includes(site)) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'id + site required' }) };
    return;
  }

  const row = await poGetRequest(site, id);
  if (!row) {
    context.res = { status: 404, headers: cors, body: JSON.stringify({ error: 'request not found' }) };
    return;
  }

  // Only the original requester can poll status (or unauth pre-auth-flag flip).
  if (process.env.AUTH_REQUIRED === '1' && row.requestedBy !== username) {
    await poAuditLog(context, { action: 'status-deny', user: username, po: row.po, site, success: false, requestId: id, message: 'not requester' });
    context.res = { status: 403, headers: cors, body: JSON.stringify({ error: 'forbidden' }) };
    return;
  }

  if (row.status === 'ready') {
    const { account: poAcc, key: poKey } = getPoBlobsStorage();
    const sasUrl = generateBlobSas(poAcc, poKey, PO_BLOB_CONTAINER, row.blobPath, 'r', PO_USER_SAS_TTL_MIN * 60);
    await poAuditLog(context, { action: 'sas-issued', user: username, po: row.po, site, success: true, requestId: id, message: row.blobPath });
    context.res = { status: 200, headers: cors, body: JSON.stringify({ status: 'ready', sasUrl, blobPath: row.blobPath, expiresInSec: PO_USER_SAS_TTL_MIN * 60 }) };
    return;
  }

  context.res = { status: 200, headers: cors, body: JSON.stringify({ status: row.status, error: row.error || undefined }) };
}

// ----- Agent: poll for next request ------------------------------------------

async function handlePoAgentPoll(context, req, cors) {
  const site = String((req.query && req.query.site) || '').toUpperCase();
  const agentUser = (req.authUser && req.authUser.username) || 'agent';

  if (!PO_VALID_SITES.includes(site)) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'site required' }) };
    return;
  }

  const claimed = await poClaimNextPending(site, agentUser);
  if (!claimed) {
    context.res = { status: 200, headers: cors, body: JSON.stringify({ empty: true }) };
    return;
  }

  await poAuditLog(context, { action: 'agent-claim', user: agentUser, po: claimed.po, site, success: true, requestId: claimed.id });
  context.res = { status: 200, headers: cors, body: JSON.stringify({ id: claimed.id, po: claimed.po, site }) };
}

// ----- Agent: get a per-blob write SAS ---------------------------------------

async function handlePoAgentWriteSas(context, req, cors) {
  const id = (req.query && req.query.id) || '';
  const site = String((req.query && req.query.site) || '').toUpperCase();
  const blobPath = (req.query && req.query.blobPath) || '';
  const agentUser = (req.authUser && req.authUser.username) || 'agent';

  if (!id || !PO_VALID_SITES.includes(site) || !blobPath) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'id + site + blobPath required' }) };
    return;
  }
  // Reject any path traversal attempt + enforce site-prefix scoping.
  if (blobPath.includes('..') || blobPath.startsWith('/') || blobPath.includes('\\')) {
    await poAuditLog(context, { action: 'agent-sas-deny', user: agentUser, po: '', site, success: false, requestId: id, message: 'bad path: ' + blobPath });
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'invalid blobPath' }) };
    return;
  }
  if (!blobPath.startsWith(site + '/')) {
    await poAuditLog(context, { action: 'agent-sas-deny', user: agentUser, po: '', site, success: false, requestId: id, message: 'site-prefix mismatch: ' + blobPath });
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'blobPath must start with ' + site + '/' }) };
    return;
  }

  // Verify the request is actually claimed (not pending, not done).
  const row = await poGetRequest(site, id);
  if (!row || row.status !== 'in-progress') {
    await poAuditLog(context, { action: 'agent-sas-deny', user: agentUser, po: row ? row.po : '', site, success: false, requestId: id, message: 'request not in-progress' });
    context.res = { status: 409, headers: cors, body: JSON.stringify({ error: 'request not claimed by an agent' }) };
    return;
  }

  const { account: poAcc, key: poKey } = getPoBlobsStorage();
  const writeSas = generateBlobSas(poAcc, poKey, PO_BLOB_CONTAINER, blobPath, 'cw', PO_AGENT_SAS_TTL_SEC);
  await poAuditLog(context, { action: 'agent-sas-issued', user: agentUser, po: row.po, site, success: true, requestId: id, message: blobPath });
  context.res = { status: 200, headers: cors, body: JSON.stringify({ writeSas, blobPath, expiresInSec: PO_AGENT_SAS_TTL_SEC }) };
}

// ----- Agent: report completion ----------------------------------------------

async function handlePoAgentComplete(context, req, cors) {
  if (req.method !== 'POST') {
    context.res = { status: 405, headers: cors, body: JSON.stringify({ error: 'POST required' }) };
    return;
  }
  const body = req.body || {};
  const id = String(body.id || '');
  const site = String(body.site || '').toUpperCase();
  const ok = !!body.ok;
  const blobPath = String(body.blobPath || '');
  const error = String(body.error || '');
  const agentUser = (req.authUser && req.authUser.username) || 'agent';

  if (!id || !PO_VALID_SITES.includes(site)) {
    context.res = { status: 400, headers: cors, body: JSON.stringify({ error: 'id + site required' }) };
    return;
  }

  const row = await poGetRequest(site, id);
  if (!row) {
    context.res = { status: 404, headers: cors, body: JSON.stringify({ error: 'request not found' }) };
    return;
  }

  const { account, key } = getStorage();
  const updated = Object.assign({}, row, {
    status: ok ? 'ready' : 'failed',
    blobPath: ok ? blobPath : '',
    error: ok ? '' : (error || 'agent reported failure'),
    updatedAt: new Date().toISOString()
  });
  await tableReq(account, key, 'PUT',
    `/popdfrequests(PartitionKey='${encodeURIComponent(site)}',RowKey='${encodeURIComponent(id)}')`,
    sanitizeForTable(updated), { 'If-Match': '*' });

  await poAuditLog(context, { action: ok ? 'agent-complete-ok' : 'agent-complete-fail', user: agentUser, po: row.po, site, success: ok, requestId: id, message: ok ? blobPath : error });
  context.res = { status: 200, headers: cors, body: JSON.stringify({ ok: true }) };
}

// ══════════════════════════════════════════
// PO PDF HELPERS
// ══════════════════════════════════════════

function getPoBlobsStorage() {
  const account = process.env.POBLOBS_ACCOUNT_NAME;
  const key = process.env.POBLOBS_ACCOUNT_KEY;
  if (!account || !key) throw new Error('POBLOBS_ACCOUNT_NAME/KEY not configured');
  return { account, key };
}

function clientIp(req) {
  // Azure Functions sets X-Forwarded-For; first hop is the actual client.
  const h = req.headers || {};
  const xff = h['x-forwarded-for'] || h['X-Forwarded-For'] || '';
  if (xff) return String(xff).split(',')[0].trim().split(':')[0];
  return h['x-azure-clientip'] || '';
}

// Service SAS for one specific blob. Permissions: 'r' (read) or 'cw' (create+write).
// canonicalizedResource uses the unencoded path; URL path uses encoded segments.
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

async function poGetRequest(site, id) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'popdfrequests');
  try {
    const raw = await tableReq(account, key, 'GET',
      `/popdfrequests(PartitionKey='${encodeURIComponent(site)}',RowKey='${encodeURIComponent(id)}')`);
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

// Atomically claim the oldest pending request for a site. Returns null if none.
// Also re-claims any in-progress row whose claim has expired (agent crash).
async function poClaimNextPending(site, agentUser) {
  const { account, key } = getStorage();
  await ensureTable(account, key, 'popdfrequests');

  const cutoffIso = new Date(Date.now() - PO_AGENT_CLAIM_TIMEOUT_SEC * 1000).toISOString();
  const filterRaw = `PartitionKey eq '${site}' and (status eq 'pending' or (status eq 'in-progress' and claimedAt lt '${cutoffIso}'))`;
  const filter = encodeURIComponent(filterRaw);
  const url = `/popdfrequests()?$filter=${filter}&$top=10`;
  const raw = await tableReq(account, key, 'GET', url);
  const data = JSON.parse(raw);
  const rows = (data.value || []).sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));

  for (const row of rows) {
    const claim = Object.assign({}, row, {
      status: 'in-progress',
      claimedAt: new Date().toISOString(),
      claimedBy: agentUser,
      updatedAt: new Date().toISOString()
    });
    try {
      // ETag from odata is in row['odata.etag'] or row.etag; with nometadata accept it isn't returned.
      // Use unconditional PUT — the worst case is two agents both think they claimed; the second
      // will overwrite the first's claimedBy but the work itself is idempotent (same blob path,
      // same final SAS), so impact is just a wasted upload, no correctness issue.
      await tableReq(account, key, 'PUT',
        `/popdfrequests(PartitionKey='${encodeURIComponent(site)}',RowKey='${encodeURIComponent(row.RowKey)}')`,
        sanitizeForTable(claim), { 'If-Match': '*' });
      return { id: row.RowKey, po: row.po };
    } catch (e) {
      // Lost the race, try the next one.
      continue;
    }
  }
  return null;
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
      requestId: evt.requestId || '',
      at: now.toISOString()
    };
    await tableReq(account, key, 'POST', '/poaudit', entity);
  } catch (e) {
    context.log.warn('poAuditLog failed:', e.message);
  }
}

// Optimistic-concurrency rate limit: PK=user, RK=YYYY-MM-DD-HH, count++.
// Up to 5 retries on If-Match conflict; gives best-effort enforcement.
async function poRateLimitCheck(context, user) {
  if (!user || user === 'unauth') return false; // pre-auth-flag, don't gate
  const { account, key } = getStorage();
  await ensureTable(account, key, 'poratelimit');
  const now = new Date();
  const bucket = now.toISOString().slice(0, 13); // YYYY-MM-DDTHH

  for (let attempt = 0; attempt < 5; attempt++) {
    let row = null;
    try {
      const raw = await tableReq(account, key, 'GET',
        `/poratelimit(PartitionKey='${encodeURIComponent(user)}',RowKey='${encodeURIComponent(bucket)}')`);
      row = JSON.parse(raw);
    } catch (_) { /* not found → first request this bucket */ }

    const count = row ? (row.count || 0) : 0;
    if (count >= PO_RATE_LIMIT_PER_HOUR) return true;

    const next = {
      PartitionKey: user,
      RowKey: bucket,
      count: count + 1,
      updatedAt: now.toISOString()
    };
    try {
      if (row) {
        await tableReq(account, key, 'PUT',
          `/poratelimit(PartitionKey='${encodeURIComponent(user)}',RowKey='${encodeURIComponent(bucket)}')`,
          next, { 'If-Match': '*' });
      } else {
        await tableReq(account, key, 'POST', '/poratelimit', next);
      }
      return false;
    } catch (e) {
      // 409 (already exists) or 412 (etag) — retry the read+write loop
      continue;
    }
  }
  // Fallthrough: too much contention → fail open (don't block legitimate users)
  context.log.warn(`poRateLimitCheck contention exceeded for ${user}`);
  return false;
}

// Strip Table-incompatible fields (odata metadata, Timestamp) before PUT.
function sanitizeForTable(row) {
  const out = {};
  for (const k of Object.keys(row)) {
    if (k.startsWith('odata.') || k === 'Timestamp' || k === 'etag') continue;
    out[k] = row[k];
  }
  return out;
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
