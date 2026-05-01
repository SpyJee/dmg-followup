const https = require('https');
const crypto = require('crypto');

// ══════════════════════════════════════════
// PO PDF cleanup — TimerTrigger, runs every 15min.
// Deletes any blob in dmgpdfvault/po-pdfs older than PO_BLOB_TTL_MIN (default 60).
// Lifecycle policy on the storage account is a 24h backstop; this is the real
// one-hour kill that enforces "PO files only visit cloud, never live there"
// per the PIPEDA/Loi 25 retention-minimization design.
// ══════════════════════════════════════════

const CONTAINER = process.env.PO_BLOB_CONTAINER || 'po-pdfs';
const TTL_MIN = parseInt(process.env.PO_BLOB_TTL_MIN || '60', 10);

module.exports = async function (context, timer) {
  const account = process.env.POBLOBS_ACCOUNT_NAME;
  const key = process.env.POBLOBS_ACCOUNT_KEY;
  if (!account || !key) {
    context.log.error('POBLOBS_ACCOUNT_NAME/KEY missing — skipping cleanup');
    return;
  }

  const cutoff = new Date(Date.now() - TTL_MIN * 60 * 1000);
  context.log(`po-cleanup: removing blobs in ${account}/${CONTAINER} older than ${cutoff.toISOString()} (TTL=${TTL_MIN}min)`);

  let scanned = 0;
  let deleted = 0;
  let failed = 0;
  let marker = null;

  while (true) {
    const qs = new URLSearchParams({ restype: 'container', comp: 'list', maxresults: '5000' });
    if (marker) qs.set('marker', marker);
    let xml;
    try {
      xml = await blobReq(account, key, 'GET', `/${CONTAINER}?${qs.toString()}`);
    } catch (e) {
      context.log.error('list failed:', e.message);
      return;
    }

    // Lightweight regex parse — we only need <Name>...</Name> and <Last-Modified>...</Last-Modified>.
    const blobBlocks = xml.match(/<Blob>[\s\S]*?<\/Blob>/g) || [];
    for (const b of blobBlocks) {
      scanned++;
      const name = (b.match(/<Name>([\s\S]*?)<\/Name>/) || [])[1];
      const lm = (b.match(/<Last-Modified>([\s\S]*?)<\/Last-Modified>/) || [])[1];
      if (!name || !lm) continue;
      const lmDate = new Date(lm);
      if (lmDate >= cutoff) continue;
      try {
        const encoded = name.split('/').map(encodeURIComponent).join('/');
        await blobReq(account, key, 'DELETE', `/${CONTAINER}/${encoded}`);
        deleted++;
      } catch (e) {
        failed++;
        context.log.warn(`delete failed for ${name}:`, e.message);
      }
    }

    const nextMarker = (xml.match(/<NextMarker>([\s\S]*?)<\/NextMarker>/) || [])[1];
    if (!nextMarker || !nextMarker.trim()) break;
    marker = nextMarker.trim();
  }

  context.log(`po-cleanup: scanned=${scanned} deleted=${deleted} failed=${failed}`);
};

// ----- Blob SharedKey HMAC signer (no SDK) -----------------------------------

function blobReq(account, key, method, pathWithQuery) {
  return new Promise((resolve, reject) => {
    const date = new Date().toUTCString();
    const [pathOnly, qs = ''] = pathWithQuery.split('?');
    const canonicalizedHeaders = `x-ms-date:${date}\nx-ms-version:2020-08-04\n`;
    const canonicalizedResource = `/${account}${pathOnly}` + canonicalizeQuery(qs, account, pathOnly);
    // Spec: VERB \n CE \n CL \n CLEN \n MD5 \n CT \n DATE \n IMS \n IM \n INM \n IUS \n RANGE \n
    // CHEADERS + CRESOURCE.   13 fields total, all empty here except VERB and the
    // tail (we use x-ms-date inside canonicalizedHeaders, so the Date slot stays "").
    const stringToSign = [
      method, '', '', '', '', '', '', '', '', '', '', '',
      canonicalizedHeaders + canonicalizedResource
    ].join('\n');

    const sig = crypto.createHmac('sha256', Buffer.from(key, 'base64'))
      .update(stringToSign, 'utf8').digest('base64');

    const opts = {
      hostname: `${account}.blob.core.windows.net`,
      path: pathWithQuery,
      method,
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

function canonicalizeQuery(qs) {
  if (!qs) return '';
  const params = new URLSearchParams(qs);
  const map = {};
  for (const [k, v] of params.entries()) {
    const lk = k.toLowerCase();
    map[lk] = (map[lk] ? map[lk] + ',' : '') + v;
  }
  const keys = Object.keys(map).sort();
  return keys.map(k => `\n${k}:${map[k]}`).join('');
}
