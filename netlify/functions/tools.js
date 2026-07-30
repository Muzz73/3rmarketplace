/* 3R Tool Pool — members-only equipment feed.
   Reads a published Google Sheet (CSV) whose URL is stored in the
   TOOLS_SHEET_URL environment variable (kept server-side, never sent to the
   browser), and returns the listings as JSON ONLY to a logged-in partner.
   Anyone not signed in gets 401 and no data. */

exports.handler = async (event, context) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };

  // Require a valid Netlify Identity user. When the site is logged in it sends
  // the Identity token, and Netlify decodes it into context.clientContext.user.
  const user = context && context.clientContext && context.clientContext.user;
  if (!user) {
    return { statusCode: 401, headers, body: JSON.stringify({ error: 'Not authorised' }) };
  }

  const url = process.env.TOOLS_SHEET_URL;
  if (!url) {
    // Not configured yet — return empty so the site shows a friendly message.
    return { statusCode: 200, headers, body: JSON.stringify({ tools: [], note: 'sheet-not-configured' }) };
  }

  try {
    const res = await fetch(url, { redirect: 'follow' });
    if (!res.ok) throw new Error('sheet ' + res.status);
    const csv = await res.text();
    return { statusCode: 200, headers, body: JSON.stringify({ tools: parseTools(csv) }) };
  } catch (e) {
    return { statusCode: 502, headers, body: JSON.stringify({ error: 'Could not load listings' }) };
  }
};

/* ---- CSV -> tool objects, matching the sheet columns by header name ---- */
function parseTools(text) {
  const rows = csvRows(text);
  if (rows.length < 2) return [];
  const head = rows[0].map((h) => String(h).trim().toLowerCase());
  const col = (names) => { for (const n of names) { const i = head.indexOf(n); if (i !== -1) return i; } return -1; };
  const iName = col(['name', 'equipment', 'item']);
  const iCat = col(['category', 'cat']);
  const iLoc = col(['location', 'loc']);
  const iDesc = col(['description', 'desc', 'details']);
  const iCond = col(['condition', 'cond']);
  const iRate = col(['day rate', 'daily rate', 'rate', 'price']);
  const iDate = col(['date listed', 'date', 'listed']);
  const iStatus = col(['status']);
  const iPhoto = col(['photo url', 'photo', 'image url', 'image']);
  const iRef = col(['ref', 'reference', 'id']);

  const cell = (row, i) => (i >= 0 && row[i] != null ? String(row[i]).trim() : '');
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const name = cell(row, iName);
    if (!name) continue; // skip blank rows
    const status = cell(row, iStatus).toLowerCase();
    if (status && status !== 'available') continue; // only Available (or blank) shows
    out.push({
      name: name,
      cat: cell(row, iCat) || 'Other',
      cond: cell(row, iCond),
      loc: cell(row, iLoc),
      desc: cell(row, iDesc),
      rate: cell(row, iRate).replace(/[^0-9.]/g, ''),
      date: cell(row, iDate),
      photo: cell(row, iPhoto),
      ref: cell(row, iRef)
    });
  }
  return out;
}

/* Minimal CSV parser: handles quoted fields, embedded commas and newlines. */
function csvRows(text) {
  const rows = []; let row = []; let field = ''; let inQ = false;
  text = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else { inQ = false; } }
      else { field += c; }
    } else if (c === '"') { inQ = true; }
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
    else { field += c; }
  }
  row.push(field); rows.push(row);
  return rows;
}
