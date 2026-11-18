/**
 * Gmail Cleanup – audit. READ-ONLY: nothing in Gmail is changed, moved or deleted.
 *
 * Reads the headers of every message (Gmail itself leaves out spam and trash) and writes:
 *   Raw data  one row per message (metadata only, never the body)
 *   Senders   overview per sender, sorted by message count
 *   Domains   overview per domain
 *   Status    progress
 *
 * Apps Script runs at most 6 minutes at a time, so the audit works in chunks and starts
 * the next chunk itself through a trigger until everything has been read.
 */

function startAudit() {
  stopAudit();
  const props = PropertiesService.getScriptProperties();
  ['pageToken', 'count', 'done', 'failedRuns', 'started', 'total'].forEach(k => props.deleteProperty(k));
  props.setProperties({
    started: new Date().toISOString(),
    total: String(Gmail.Users.getProfile('me').messagesTotal || '')
  });
  const raw = resetSheet_(SHEETS.RAW, RAW_COLS);
  raw.getRange(2, 1, raw.getMaxRows() - 1, 2).setNumberFormat('@');   // IDs as text, not numbers
  resetSheet_(SHEETS.ERRORS, ['ID', 'Error']);
  setAuditStatus_('Running', 0);
  runBatch();
}

function stopAudit() {
  removeHandlerTriggers_('runBatch');
  if (PropertiesService.getScriptProperties().getProperty('started')) setAuditStatus_('Stopped');
}

function runBatch() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;   // another run is in progress
  const props = PropertiesService.getScriptProperties();
  try {
    if (props.getProperty('done') === '1') return;
    const t0 = Date.now();
    const labelNames = loadLabelNames_();
    const raw = sheet_(SHEETS.RAW);
    let count = Number(props.getProperty('count') || 0);

    while (Date.now() - t0 < CFG.MAX_RUN_MS) {
      const opts = { maxResults: CFG.PAGE_SIZE };
      const token = props.getProperty('pageToken');
      if (token) opts.pageToken = token;

      const page = withRetry_(() => Gmail.Users.Messages.list('me', opts));
      const ids = (page.messages || []).map(m => m.id);
      const rows = fetchMetadata_(ids).map(m => toRow_(m, labelNames));
      if (rows.length) raw.getRange(raw.getLastRow() + 1, 1, rows.length, RAW_COLS.length).setValues(rows);
      count += rows.length;
      props.setProperties({ count: String(count), failedRuns: '0' });

      if (!page.nextPageToken) {
        props.setProperty('done', '1');
        props.deleteProperty('pageToken');
        removeHandlerTriggers_('runBatch');
        setAuditStatus_('Retrying failed messages …', count);
        count += retryFailed_(labelNames);
        props.setProperty('count', String(count));
        setAuditStatus_('Read complete, building overview …', count);
        buildOverview();
        setAuditStatus_('Done', count);
        return;
      }
      props.setProperty('pageToken', page.nextPageToken);
      setAuditStatus_('Running', count);
    }
    scheduleHandler_('runBatch');
  } catch (e) {
    const failed = Number(props.getProperty('failedRuns') || 0) + 1;
    props.setProperty('failedRuns', String(failed));
    if (failed < CFG.MAX_FAILED_RUNS) {
      setAuditStatus_('Error (retrying in 1 min): ' + e.message);
      scheduleHandler_('runBatch');
    } else {
      removeHandlerTriggers_('runBatch');
      setAuditStatus_('Stopped after repeated errors: ' + e.message);
    }
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** Tries the messages that failed during the run once more. Returns how many were read. */
function retryFailed_(labelNames) {
  const sh = sheet_(SHEETS.ERRORS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  resetSheet_(SHEETS.ERRORS, ['ID', 'Error']);   // fetchMetadata_ writes the ones that still fail
  const rows = fetchMetadata_(ids).map(m => toRow_(m, labelNames));
  appendRows_(sheet_(SHEETS.RAW), rows);
  return rows.length;
}

// ---------- Overview ----------

function buildOverview() {
  const raw = sheet_(SHEETS.RAW);
  const n = raw.getLastRow() - 1;
  const data = n > 0 ? raw.getRange(2, 1, n, RAW_COLS.length).getValues() : [];

  // Drop duplicates (if a run was interrupted in the middle of a page)
  const seen = new Set();
  const msgs = data.filter(r => !seen.has(r[0]) && seen.add(r[0]));

  // Threads you have written in = conversations you took part in
  const repliedThreads = new Set(msgs.filter(r => String(r[8]).split(', ').includes('SENT')).map(r => r[1]));

  const senders = {};
  for (const r of msgs) {
    const [, thread, date, email, name, domain, subject, bytes, labelsStr, unsub] = r;
    const labels = String(labelsStr).split(', ');
    if (labels.includes('SENT')) continue;   // your own outgoing mail

    const s = senders[email] || (senders[email] = {
      email, name, domain, n: 0, inbox: 0, unread: 0, starred: 0, important: 0,
      replied: 0, cats: {}, unsub: '', bytes: 0, first: date, last: date, subjects: []
    });
    s.n++;
    s.bytes += Number(bytes) || 0;
    if (!s.name && name) s.name = name;
    if (labels.includes('INBOX')) s.inbox++;
    if (labels.includes('UNREAD')) s.unread++;
    if (labels.includes('STARRED')) s.starred++;
    if (labels.includes('IMPORTANT')) s.important++;
    if (repliedThreads.has(thread)) s.replied++;
    labels.forEach(l => { if (CATEGORY_NAMES[l]) s.cats[l] = (s.cats[l] || 0) + 1; });
    if (UNSUB_RANK[unsub] > UNSUB_RANK[s.unsub]) s.unsub = unsub;
    if (date < s.first) s.first = date;
    if (date > s.last) s.last = date;
    if (s.subjects.length < 3 && subject && !s.subjects.includes(subject)) s.subjects.push(subject);
  }

  const list = Object.values(senders).sort((a, b) => b.n - a.n);
  const pct = (part, whole) => whole ? Math.round(100 * part / whole) : 0;
  const mb = bytes => Math.round(bytes / 1048576 * 10) / 10;

  writeTable_(SHEETS.SENDERS, SENDER_COLS, list.map(s => {
    const topCat = Object.keys(s.cats).sort((a, b) => s.cats[b] - s.cats[a])[0];
    return [safe_(s.email), safe_(s.name), s.domain, s.n, s.inbox, s.unread, pct(s.n - s.unread, s.n),
      s.starred, s.important, s.replied, CATEGORY_NAMES[topCat] || '', s.unsub, mb(s.bytes),
      s.first, s.last, safe_(s.subjects.join(' | '))];
  }), { dateCols: [14, 15] });

  const domains = {};
  for (const s of list) {
    const d = domains[s.domain] || (domains[s.domain] = {
      domain: s.domain, senders: 0, n: 0, unread: 0, replied: 0, bytes: 0, last: s.last });
    d.senders++; d.n += s.n; d.unread += s.unread; d.replied += s.replied; d.bytes += s.bytes;
    if (s.last > d.last) d.last = s.last;
  }
  writeTable_(SHEETS.DOMAINS, DOMAIN_COLS, Object.values(domains).sort((a, b) => b.n - a.n).map(d =>
    [d.domain, d.senders, d.n, pct(d.n - d.unread, d.n), d.replied, mb(d.bytes), d.last]), { dateCols: [7] });
}

/**
 * A compact copy of "Senders" as the FIRST sheet, so it can be exported as CSV (Drive exports
 * only the first sheet) – e.g. to review the senders with an AI assistant.
 */
function exportCompact() {
  const src = sheet_(SHEETS.SENDERS);
  const n = src.getLastRow() - 1;
  const data = n > 0 ? src.getRange(2, 1, n, SENDER_COLS.length).getValues() : [];
  const year = d => (d instanceof Date ? d.getFullYear() : '');
  const rows = data.filter(r => Number(r[3]) >= CFG.COMPACT_MIN_COUNT).map(r => [
    safe_(r[0]), safe_(String(r[1]).slice(0, 40)), r[3], r[6], r[7], r[9], r[10], r[11], r[12],
    year(r[13]), year(r[14]),
    safe_(String(r[15]).split(' | ').slice(0, 2).map(s => s.trim().slice(0, 70)).join(' | '))
  ]);
  writeTable_(SHEETS.COMPACT, COMPACT_COLS, rows);
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(sheet_(SHEETS.COMPACT));
  ss.moveActiveSheet(1);
}

function setAuditStatus_(state, count) {
  const props = PropertiesService.getScriptProperties();
  const read = count !== undefined ? count : Number(props.getProperty('count') || 0);
  sheet_(SHEETS.STATUS).getRange(1, 1, 5, 2).setValues([
    ['Audit', state],
    ['Messages read', read],
    ['Messages in total (approx., incl. sent)', props.getProperty('total') || ''],
    ['Started', props.getProperty('started') || ''],
    ['Last updated', new Date()]
  ]);
}
