const https = require('https');
const crypto = require('crypto');

// ══════════════════════════════════════════
// DMG Aerospace — Production Follow-Up API
// Routes:
//   POST /api/followup/{order-save}      — Save order
//   GET  /api/followup/{order-list}      — List all orders
//   GET  /api/followup/{order-get}       — Get single order
//   POST /api/followup/{order-delete}    — Delete order
//   POST /api/followup/{recipe-save}     — Save recipe
//   GET  /api/followup/{recipe-list}     — List all recipes
//   POST /api/followup/{recipe-delete}   — Delete recipe
//   POST /api/followup/{suivi-update}    — Update operation status
//   GET  /api/followup/{suivi-get}       — Get all statuses for an order
//   GET  /api/followup/{holidays}        — Get holiday list
//   GET  /api/followup/{health}          — Health check
// ══════════════════════════════════════════

module.exports = async function (context, req) {
  const cors = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  if (req.method === 'OPTIONS') {
    context.res = { status: 204, headers: cors, body: '' };
    return;
  }

  const route = req.params.route || '';
  context.log(`Follow-Up Function: ${req.method} /${route}`);

  try {
    if (route === 'health') {
      context.res = { status: 200, headers: cors, body: JSON.stringify({ status: 'ok', timestamp: new Date().toISOString() }) };
      return;
    }

    // Routes will be implemented in Phase 2
    context.res = { status: 404, headers: cors, body: JSON.stringify({ error: 'Unknown route: ' + route }) };
  } catch (err) {
    context.log.error('Handler error:', err.message, err.stack);
    context.res = { status: 500, headers: cors, body: JSON.stringify({ error: err.message }) };
  }
};

// ── TABLE STORAGE HELPERS (reused from dmg-quoting) ──

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
