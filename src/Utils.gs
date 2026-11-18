/**
 * Gmail Cleanup – shared helpers: sheets, Gmail calls, parsing and triggers.
 */

// ---------- Sheets ----------

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function resetSheet_(name, header) {
  const sh = sheet_(name);
  const filter = sh.getFilter();
  if (filter) filter.remove();
  sh.clear();
  if (header.length) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Replaces a sheet with a header and rows. textCols are formatted as text before writing,
 * so Gmail IDs that look like numbers are not converted by Sheets.
 */
function writeTable_(name, header, rows, opts) {
  opts = opts || {};
  const sh = resetSheet_(name, header);
  if (!rows.length) return sh;
  (opts.textCols || []).forEach(c => sh.getRange(2, c, rows.length, 1).setNumberFormat('@'));
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  (opts.dateCols || []).forEach(c => sh.getRange(2, c, rows.length, 1).setNumberFormat('yyyy-mm-dd'));
  sh.getRange(1, 1, rows.length + 1, header.length).createFilter();
  return sh;
}

function appendRows_(sh, rows, header) {
  if (header && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (!rows.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function setValidation_(sh, col, values) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true).build();
  sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

/** Stops text such as "=…" or "+…" from being read as a formula. */
function safe_(s) {
  s = String(s || '').slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// ---------- Gmail ----------

/** Runs fn again after a pause when Gmail answers with a quota or temporary error. */
function withRetry_(fn) {
  for (let attempt = 0; ; attempt++) {
    try {
      return fn();
    } catch (e) {
      if (attempt >= 3 || !/quota|rate|limit|backend|internal/i.test(e.message)) throw e;
      Utilities.sleep(20000 * (attempt + 1));
    }
  }
}

function listIds_(q) {
  const out = [];
  let pageToken;
  do {
    const opts = { maxResults: CFG.PAGE_SIZE };
    if (q) opts.q = q;
    if (pageToken) opts.pageToken = pageToken;
    const page = withRetry_(() => Gmail.Users.Messages.list('me', opts));
    (page.messages || []).forEach(m => out.push(m));
    pageToken = page.nextPageToken;
  } while (pageToken);
  return out;
}

/**
 * Fetches header metadata one message at a time through the Gmail advanced service, paced
 * below the per-minute quota. (The advanced service does not count against the daily
 * UrlFetch quota, which a parallel version quickly exhausts.) Failures go to "Errors".
 */
function fetchMetadata_(ids) {
  const minMs = 60000 / CFG.MAX_PER_MINUTE;
  const out = [];
  let todo = ids.slice();
  let lastError = '';

  for (let attempt = 0; attempt < 4 && todo.length; attempt++) {
    if (attempt) Utilities.sleep(10000 * attempt);   // the quota is per minute, so wait a while
    const retry = [];
    for (const id of todo) {
      const start = Date.now();
      try {
        out.push(Gmail.Users.Messages.get('me', id, { format: 'metadata', metadataHeaders: META_HEADERS }));
      } catch (e) {
        if (!/not found/i.test(e.message)) {   // "not found" = deleted in the meantime
          retry.push(id);
          lastError = e.message.slice(0, 300);
        }
      }
      const wait = minMs - (Date.now() - start);
      if (wait > 0) Utilities.sleep(wait);
    }
    todo = retry;
  }

  if (todo.length) appendRows_(sheet_(SHEETS.ERRORS), todo.map(id => [id, lastError]), ['ID', 'Error']);
  return out;
}

function headerMap_(m) {
  const h = {};
  ((m.payload && m.payload.headers) || []).forEach(x => { h[x.name.toLowerCase()] = x.value; });
  return h;
}

function parseFrom_(s) {
  const m = String(s || '').match(/^\s*"?([^"<]*?)"?\s*<([^>]+)>/);
  if (m) return { name: m[1].trim(), email: m[2].trim().toLowerCase() };
  return { name: '', email: String(s || '').trim().toLowerCase() };
}

/** How a sender can be unsubscribed: one-click (RFC 8058), link, mail – or not at all. */
function unsubMethod_(h) {
  const v = h['list-unsubscribe'] || '';
  if (!v) return '';
  const hasHttp = /<https?:/i.test(v);
  if (hasHttp && /one-click/i.test(h['list-unsubscribe-post'] || '')) return 'one-click';
  if (hasHttp) return 'link';
  if (/<mailto:/i.test(v)) return 'mail';
  return '';
}

function loadLabelNames_() {
  const names = {};
  (Gmail.Users.Labels.list('me').labels || []).forEach(l => { names[l.id] = l.name; });
  return names;
}

/** One message as a sheet row (see RAW_COLS). */
function toRow_(m, labelNames) {
  const h = headerMap_(m);
  const from = parseFrom_(h['from']);
  const labels = (m.labelIds || []).map(id => labelNames[id] || id);
  return [
    m.id, m.threadId, new Date(Number(m.internalDate)), safe_(from.email), safe_(from.name),
    from.email.split('@')[1] || '', safe_(h['subject'] || ''), m.sizeEstimate || 0,
    labels.join(', '), unsubMethod_(h)
  ];
}

// ---------- Triggers and status ----------

function scheduleHandler_(name) {
  removeHandlerTriggers_(name);
  ScriptApp.newTrigger(name).timeBased().after(60 * 1000).create();
}

function removeHandlerTriggers_(name) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === name)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function daysAgo_(days) {
  return new Date(Date.now() - Number(days) * 86400000);
}

/** Compiles a regex from a setting; an empty setting matches nothing. */
function regex_(source) {
  return source ? new RegExp(String(source), 'i') : null;
}

function matches_(re, text) {
  return !!re && re.test(String(text || ''));
}
