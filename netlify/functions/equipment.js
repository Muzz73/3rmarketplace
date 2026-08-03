/* 3R Marketplace — public equipment feed (Browse Equipment).

   Reads the Smart Master workbook from 3R SharePoint via Microsoft Graph,
   server-side, and returns ONLY:
     • rows whose Status = "Available"
     • the public columns (never Part/Serial No, Location, Owner, Notes, Status)

   The client secret lives in Netlify environment variables and never reaches
   the browser. Columns are matched BY HEADING NAME, so the sheet's columns can
   be reordered freely without breaking the site.

   Until the Graph credentials are configured this serves a bundled snapshot
   (equipment-sample.json) so the Browse page can be built and reviewed first.

   Environment variables (Netlify ▸ Site settings ▸ Environment variables):
     MS_TENANT_ID      Directory (tenant) ID
     MS_CLIENT_ID      Application (client) ID
     MS_CLIENT_SECRET  client secret VALUE
     MS_DRIVE_ID       drive (document library) id
     MS_ITEM_ID        driveItem id of the workbook
     EQUIP_TABLE       optional, defaults to "Equipment"
     EQUIP_IMG_BASE    optional, defaults to "/equipment-images/"
*/

const SAMPLE = require('./equipment-sample.json');

const TABLE = process.env.EQUIP_TABLE || 'Equipment';
const IMG_BASE = process.env.EQUIP_IMG_BASE || '/equipment-images/';
const CACHE_MS = 5 * 60 * 1000; // 5 minutes

// Public column heading -> output key. Anything not listed is never published.
const PUBLIC_COLUMNS = {
  'item': 'item',
  'category': 'cat',
  'sub-category': 'sub',
  'description': 'name',
  'quantity': 'qty',
  'original manufacturer': 'oem',
  'condition': 'cond',
  'documentation': 'doc',
  'price': 'price',
  'details': 'details',
  'image': 'img'
};

// Warm-container cache. Also acts as the fallback if Graph is unavailable.
let cache = { at: 0, payload: null };

exports.handler = async (event) => {
  const headers = {
    'Content-Type': 'application/json',
    'Cache-Control': 'public, max-age=300'
  };
  const fresh = !!(event && event.queryStringParameters && event.queryStringParameters.fresh);

  // Serve from cache unless a refresh was requested.
  if (!fresh && cache.payload && Date.now() - cache.at < CACHE_MS) {
    return ok(headers, Object.assign({}, cache.payload, { cached: true }));
  }

  const cfg = {
    tenant: process.env.MS_TENANT_ID,
    client: process.env.MS_CLIENT_ID,
    secret: process.env.MS_CLIENT_SECRET,
    drive: process.env.MS_DRIVE_ID,
    item: process.env.MS_ITEM_ID
  };

  // Not wired to SharePoint yet — serve the bundled snapshot.
  if (!cfg.tenant || !cfg.client || !cfg.secret || !cfg.drive || !cfg.item) {
    const payload = {
      equipment: (SAMPLE.equipment || []).map(normaliseImage),
      total: (SAMPLE.equipment || []).length,
      source: 'sample',
      note: 'Graph not configured — serving bundled snapshot'
    };
    cache = { at: Date.now(), payload: payload };
    return ok(headers, payload);
  }

  try {
    const token = await getToken(cfg);
    /* Table is addressed as /workbook/tables/{name} — do NOT wrap the name in
       quotes here; quotes belong only to the tables('name') bracket form, and
       encoding them makes Graph look for a table literally called "'Equipment'"
       (404). Drive and item ids are used as-is: they already contain only
       URL-safe characters (b!… ids include "!", which Graph accepts). */
    const base =
      'https://graph.microsoft.com/v1.0/drives/' + cfg.drive +
      '/items/' + cfg.item +
      '/workbook/tables/' + encodeURIComponent(TABLE);

    const [head, rows] = await Promise.all([
      graph(base + '/headerRowRange?$select=values', token),
      graph(base + '/rows?$select=values&$top=5000', token)
    ]);

    const payload = {
      equipment: buildRows(head, rows),
      source: 'sharepoint',
      fetchedAt: new Date().toISOString()
    };
    payload.total = payload.equipment.length;
    cache = { at: Date.now(), payload: payload };
    return ok(headers, payload);
  } catch (e) {
    // Locked file (423), throttled (429), transient outage — never show an empty
    // marketplace. Serve the last good response, else the bundled snapshot.
    if (cache.payload) {
      return ok(headers, Object.assign({}, cache.payload, { stale: true, warning: String(e.message || e) }));
    }
    const payload = {
      equipment: (SAMPLE.equipment || []).map(normaliseImage),
      total: (SAMPLE.equipment || []).length,
      source: 'sample-fallback',
      warning: String(e.message || e)
    };
    return ok(headers, payload);
  }
};

function ok(headers, body) {
  return { statusCode: 200, headers: headers, body: JSON.stringify(body) };
}

/* ---------- Microsoft Graph ---------- */

async function getToken(cfg) {
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: cfg.client,
    client_secret: cfg.secret,
    scope: 'https://graph.microsoft.com/.default'
  });
  const res = await fetch(
    'https://login.microsoftonline.com/' + encodeURIComponent(cfg.tenant) + '/oauth2/v2.0/token',
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: body }
  );
  if (!res.ok) throw new Error('token ' + res.status);
  const json = await res.json();
  if (!json.access_token) throw new Error('no access_token');
  return json.access_token;
}

async function graph(url, token) {
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  if (!res.ok) {
    /* Include Graph's own message — it names the actual problem (unknown table,
       item not found, forbidden) instead of leaving us to guess from a status. */
    let detail = '';
    try {
      const body = await res.json();
      detail = (body && body.error && (body.error.message || body.error.code)) || '';
    } catch (e) {}
    throw new Error('graph ' + res.status + ' on ' + url.split('/workbook')[1] +
                    (detail ? ' — ' + detail : ''));
  }
  return res.json();
}

/* ---------- sheet -> listings ---------- */

function buildRows(headerResponse, rowsResponse) {
  const headings = ((headerResponse && headerResponse.values) || [[]])[0] || [];
  // heading (lowercased) -> column index
  const index = {};
  headings.forEach(function (h, i) {
    const key = String(h == null ? '' : h).trim().toLowerCase();
    if (key && index[key] === undefined) index[key] = i;
  });

  const statusAt = index['status'];
  const out = [];
  const rows = (rowsResponse && rowsResponse.value) || [];

  rows.forEach(function (r) {
    const cells = (r.values || [[]])[0] || [];
    const val = function (i) {
      if (i === undefined || i < 0) return '';
      const v = cells[i];
      return v == null ? '' : String(v).trim();
    };

    if (!val(index['item'])) return;                                  // skip blank rows
    if (String(val(statusAt)).toLowerCase() !== 'available') return;   // only live items

    const o = {};
    Object.keys(PUBLIC_COLUMNS).forEach(function (heading) {
      o[PUBLIC_COLUMNS[heading]] = val(index[heading]);
    });
    o.sub = o.sub.replace(/^Capital Equipment:\s*/i, '').trim();       // tidy legacy prefixes
    out.push(normaliseImage(o));
  });

  return out;
}

/* A real photo URL passes through; a picked filename becomes a site path.

   `generic` tells the front end to caption the image as illustrative. A real
   photograph of the actual item must NEVER carry that note, so the rule is:

     • an external URL                     -> real photo
     • a filename starting "3R-<number>"   -> real photo of that item
       (e.g. "3R-356 Horizontal XMT.jpg")
     • any other filename                  -> one of the shared 3R stock images
*/
function isItemPhoto(name) {
  return /^3R[-\s]?\d/i.test(String(name).trim());
}
function normaliseImage(o) {
  const v = (o.img || '').trim();
  if (!v) { o.img = ''; o.generic = false; return o; }
  if (/^https?:\/\//i.test(v)) { o.img = v; o.generic = false; return o; }
  o.img = IMG_BASE + encodeURIComponent(v);
  o.generic = !isItemPhoto(v);
  return o;
}
