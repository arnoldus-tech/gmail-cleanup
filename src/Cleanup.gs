/**
 * Gmail Cleanup – rules, dry-run plan, execution and the weekly run.
 *
 *   1. Prepare rules     Creates Settings, Sender rules and Newsletters. Never overwrites your
 *                        edits; newly found newsletters are appended.
 *   2. Build plan        Decides what happens to every message in "Raw data" and writes the
 *                        Plan and Summary sheets. Changes nothing in Gmail.
 *   3. Execute plan      Moves to trash, adds labels and unsubscribes newsletters marked
 *                        "Unsubscribe". Runs in chunks until the plan is done.
 *   4. Weekly run        Every Monday: registers new newsletters (and asks you), carries out your
 *                        decisions and cleans up messages that have just left the protection period.
 */

// ---------- 1. Prepare ----------

function prepareRules() {
  ensureTable_(SHEETS.SETTINGS, SETTING_COLS, DEFAULT_SETTINGS);
  ensureTable_(SHEETS.RULES, RULE_COLS, DEFAULT_SENDER_RULES);
  setValidation_(sheet_(SHEETS.RULES), 2, Object.values(RULE));
  const added = buildNewsletterList_();
  setCleanupStatus_('Rules prepared; ' + added + ' newsletters added');
}

function ensureTable_(name, header, rows) {
  if (SpreadsheetApp.getActive().getSheetByName(name)) return;
  writeTable_(name, header, rows);
}

/** Appends senders with an unsubscribe header that are not on the list yet. */
function buildNewsletterList_() {
  const s = readSettings_();
  const rules = readRules_();
  const sh = sheet_(SHEETS.NEWS);
  const existing = new Set(Object.keys(readNewsletters_()));
  if (sh.getLastRow() === 0) resetSheet_(SHEETS.NEWS, NEWS_COLS);

  const agg = {};
  readRaw_().forEach(m => {
    if (!m.unsub || m.labels.includes('SENT')) return;
    const a = agg[m.email] || (agg[m.email] = { email: m.email, name: m.name, n: 0, last: m.date, method: m.unsub });
    a.n++;
    if (m.date > a.last) a.last = m.date;
    if (UNSUB_RANK[m.unsub] > UNSUB_RANK[a.method]) a.method = m.unsub;
  });

  const yearAgo = daysAgo_(365);
  const rows = Object.values(agg)
    .filter(a => !existing.has(a.email) && senderRule_(a.email, rules) !== RULE.KEEP &&
      !matches_(s.re.jobSender, a.email) && !isProtectedSender_(a.email, a.name))
    .sort((a, b) => b.last - a.last)
    .map(a => [safe_(a.email), safe_(a.name), a.n, a.last, a.method,
      a.last < yearAgo ? DECISION.INACTIVE : DECISION.UNSUBSCRIBE, '']);
  appendRows_(sh, rows);
  sh.getRange(2, 4, Math.max(sh.getLastRow() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');
  setValidation_(sh, 6, Object.values(DECISION));
  return rows.length;
}

// ---------- 2. Plan (dry run) ----------

function buildPlan() {
  const ctx = context_();
  const counts = {};
  const trashBySender = {};
  const rows = [];
  let total = 0, receiptsAll = 0;

  readRaw_().forEach(m => {
    total++;
    if (!m.labels.includes('SENT') && matches_(ctx.s.re.receipt, m.subject)) receiptsAll++;
    const c = classify_(m, ctx);
    const key = c.action + '\u0000' + c.rule;
    counts[key] = (counts[key] || 0) + 1;
    if (c.action === ACTION.KEEP) return;
    if (c.action === ACTION.TRASH) trashBySender[m.email] = (trashBySender[m.email] || 0) + 1;
    rows.push([m.id, m.threadId, m.date, safe_(m.email), safe_(m.subject), c.rule,
      c.action === ACTION.TRASH ? ACTION.TRASH : 'Label: ' + c.label]);
  });

  writeTable_(SHEETS.PLAN, PLAN_COLS, rows, { textCols: [1, 2], dateCols: [3] });
  writeSummary_(total, receiptsAll, counts, trashBySender);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('planRow', '0');
  props.setProperty('execFails', '0');
  setCleanupStatus_('Plan built: ' + rows.length + ' actions. Nothing has been changed yet.');
}

function writeSummary_(total, receiptsAll, counts, trashBySender) {
  const byAction = {};
  Object.values(ACTION).forEach(a => { byAction[a] = 0; });
  const ruleRows = Object.keys(counts).map(k => {
    const [action, rule] = k.split('\u0000');
    byAction[action] += counts[k];
    return [action, rule, counts[k]];
  }).sort((a, b) => b[2] - a[2]);
  const topTrash = Object.keys(trashBySender)
    .sort((a, b) => trashBySender[b] - trashBySender[a]).slice(0, 50)
    .map(e => [safe_(e), trashBySender[e], '']);

  const out = [
    ['Plan built', new Date(), ''],
    ['Messages in total', total, ''],
    ['Moved to trash', byAction[ACTION.TRASH], ''],
    ['Labelled (kept)', byAction[ACTION.LABEL], ''],
    ['Kept unchanged', byAction[ACTION.KEEP], ''],
    ['Receipts in total (always kept)', receiptsAll, ''],
    ['', '', ''],
    ['Action', 'Rule', 'Count'],
    ...ruleRows,
    ['', '', ''],
    ['Sender (most trashed)', 'Trashed', ''],
    ...topTrash
  ];
  const sh = resetSheet_(SHEETS.SUMMARY, []);
  sh.getRange(1, 1, out.length, 3).setValues(out);
  sh.getRange(1, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  [8, 8 + ruleRows.length + 2].forEach(r => sh.getRange(r, 1, 1, 3).setFontWeight('bold'));
  sh.autoResizeColumns(1, 3);
}

// ---------- 3. Execute plan ----------

function executePlanFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const n = Math.max(sheet_(SHEETS.PLAN).getLastRow() - 1, 0);
  if (!n) {
    ui.alert('There is no plan. Run "Build plan (dry run)" first.');
    return;
  }
  const answer = ui.alert('Execute plan',
    n + ' actions: messages are moved to trash (restorable for 30 days), labels are added ' +
    'and newsletters marked "Unsubscribe" are unsubscribed. Continue?', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().setProperty('execFails', '0');
  executePlan();
}

function executePlan() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  const props = PropertiesService.getScriptProperties();
  try {
    const t0 = Date.now();
    const sh = sheet_(SHEETS.PLAN);
    const total = sh.getLastRow() - 1;
    let pos = Number(props.getProperty('planRow') || 0);

    if (pos < total) {
      const s = readSettings_();
      const prot = loadProtection_(s);
      const labelCache = {};
      while (pos < total && Date.now() - t0 < CFG.MAX_RUN_MS) {
        const n = Math.min(200, total - pos);
        const log = [];
        const toLabel = {};
        sh.getRange(pos + 2, 1, n, PLAN_COLS.length).getValues().forEach(r => {
          const [id, thread, , email, subject, rule, action] = r;
          if (!isValidId_(id)) {   // Sheets may have turned the ID into a number
            log.push(logRow_(id, email, subject, rule, 'Skipped: invalid ID in sheet'));
          } else if (action === ACTION.TRASH) {
            log.push(logRow_(id, email, subject, rule, safeTrash_(String(id), String(thread), prot)));
          } else if (String(action).indexOf('Label: ') === 0) {
            const name = String(action).slice(7);
            (toLabel[name] = toLabel[name] || []).push(String(id));
            log.push(logRow_(id, email, subject, rule, action));
          }
        });
        applyLabels_(toLabel, labelCache);
        appendRows_(sheet_(SHEETS.LOG), log, LOG_COLS);
        pos += n;
        props.setProperty('planRow', String(pos));
        setCleanupStatus_('Executing plan: ' + pos + ' of ' + total);
      }
      if (pos < total) {
        scheduleHandler_('executePlan');
        return;
      }
    }

    if (!processNewsletterDecisions_(t0, false)) {
      setCleanupStatus_('Plan executed; unsubscribing newsletters …');
      scheduleHandler_('executePlan');
      return;
    }
    removeHandlerTriggers_('executePlan');
    setCleanupStatus_('Plan executed');
  } catch (e) {
    const fails = Number(props.getProperty('execFails') || 0) + 1;
    props.setProperty('execFails', String(fails));
    if (fails < CFG.MAX_FAILED_RUNS) {
      setCleanupStatus_('Error (retrying in 1 min): ' + e.message);
      scheduleHandler_('executePlan');
    } else {
      removeHandlerTriggers_('executePlan');
      setCleanupStatus_('Stopped after repeated errors: ' + e.message);
    }
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ---------- 4. Weekly run ----------

function activateWeekly() {
  removeHandlerTriggers_('weeklyRun');
  ScriptApp.newTrigger('weeklyRun').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(CFG.WEEKLY_HOUR).create();
  setCleanupStatus_('Weekly run activated (Monday ' + CFG.WEEKLY_HOUR + '–' + (CFG.WEEKLY_HOUR + 1) + ')');
}

function deactivateWeekly() {
  removeHandlerTriggers_('weeklyRun');
  setCleanupStatus_('Weekly run stopped');
}

function weeklyRun() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const t0 = Date.now();
    const ctx = context_();
    const live = ctx.s[SETTING.WEEKLY_MODE] === MODE.EXECUTE;
    const labelNames = loadLabelNames_();

    // a) New newsletters from the past week, for you to decide on
    const newSenders = addNewNewsletters_(fetchMsgs_('newer_than:8d', labelNames), ctx);

    // b) Your decisions in the Newsletters sheet
    processNewsletterDecisions_(t0, live);

    // c) Messages that have just left the protection period (3 weeks overlap if a week is missed)
    const days = Number(ctx.s[SETTING.PROTECT_DAYS]);
    const msgs = fetchMsgs_('older_than:' + days + 'd newer_than:' + (days + 21) + 'd', labelNames);
    const log = [];
    const toLabel = {};
    let trashed = 0, labelled = 0;
    for (const m of msgs) {
      if (Date.now() - t0 > CFG.MAX_RUN_MS) break;
      const c = classify_(m, ctx);
      if (c.action === ACTION.TRASH) {
        const res = live ? safeTrash_(m.id, m.threadId, ctx.prot) : '(would trash)';
        if (!live || res === TRASHED) trashed++;
        log.push(logRow_(m.id, m.email, m.subject, c.rule, res));
      } else if (c.action === ACTION.LABEL) {
        if (live) (toLabel[c.label] = toLabel[c.label] || []).push(m.id);
        labelled++;
        log.push(logRow_(m.id, m.email, m.subject, c.rule, (live ? 'Label: ' : '(would label) ') + c.label));
      }
    }
    applyLabels_(toLabel, {});
    appendRows_(sheet_(SHEETS.LOG), log, LOG_COLS);

    const stats = trashed + ' to trash, ' + labelled + ' labelled';
    setCleanupStatus_('Weekly run' + (live ? '' : ' (log only)') + ': ' + stats + ', ' + newSenders.length + ' new newsletters');
    if (String(ctx.s[SETTING.WEEKLY_MAIL]).toUpperCase() === 'YES') sendSummaryMail_(newSenders, stats, live);
  } finally {
    lock.releaseLock();
  }
}

function addNewNewsletters_(msgs, ctx) {
  const found = {};
  msgs.forEach(m => {
    if (!m.unsub || m.labels.includes('SENT') || ctx.news[m.email] || found[m.email]) return;
    if (senderRule_(m.email, ctx.rules) === RULE.KEEP || matches_(ctx.s.re.jobSender, m.email)) return;
    if (isProtectedSender_(m.email, m.name)) return;
    found[m.email] = { email: m.email, name: m.name, date: m.date, method: m.unsub };
  });
  const list = Object.values(found);
  appendRows_(sheet_(SHEETS.NEWS), list.map(a => [safe_(a.email), safe_(a.name), 1, a.date, a.method, DECISION.PENDING, NEW_STATUS]));
  list.forEach(a => { ctx.news[a.email] = { decision: DECISION.PENDING }; });
  return list;
}

function sendSummaryMail_(newSenders, stats, live) {
  const pending = Object.values(readNewsletters_()).filter(x => x.decision === DECISION.PENDING).length;
  if (!newSenders.length && !pending) return;
  const lines = [];
  if (newSenders.length) {
    lines.push('New newsletters this week:');
    newSenders.forEach(a => lines.push('- ' + (a.name ? a.name + ' ' : '') + '<' + a.email + '>'));
    lines.push('');
  }
  lines.push('Newsletters waiting for your decision: ' + pending);
  lines.push('Set "Keep" or "Unsubscribe" in the "Newsletters" sheet:');
  lines.push(SpreadsheetApp.getActive().getUrl());
  lines.push('');
  lines.push("This week's clean-up" + (live ? '' : ' (log only)') + ': ' + stats);
  sendMail_(Gmail.Users.getProfile('me').emailAddress,
    'Gmail cleanup: ' + newSenders.length + ' new newsletters', lines.join('\r\n'));
}

// ---------- Classification ----------

function context_() {
  const s = readSettings_();
  return { s, rules: readRules_(), news: readNewsletters_(), prot: loadProtection_(s) };
}

/** The first rule that matches decides: Keep, Trash or Label. */
function classify_(m, ctx) {
  const s = ctx.s;
  const L = s.labels;
  if (isProtectedSender_(m.email, m.name)) return r_('Protected sender (Private.gs)', ACTION.KEEP);
  if (m.email === ctx.prot.me || m.labels.includes('SENT')) return r_('Own message', ACTION.KEEP);
  if (m.date > daysAgo_(s[SETTING.PROTECT_DAYS])) return r_('Within protection period', ACTION.KEEP);
  if (m.labels.includes('STARRED')) return r_('Starred', ACTION.KEEP);
  if (m.labels.some(l => isUserLabel_(l, s.ownLabels))) return r_('Has your own label', ACTION.KEEP);
  if (ctx.prot.sentThreads.has(m.threadId)) return r_('Thread you wrote in', ACTION.KEEP);
  if (matches_(s.re.receipt, m.subject)) return labelOrKeep_(m, 'Receipt', L.receipts);
  if (matches_(s.re.cases, m.subject)) return labelOrKeep_(m, 'Case words', L.cases);

  const rule = senderRule_(m.email, ctx.rules);
  if (rule === RULE.KEEP) return r_('Sender rule: Keep', ACTION.KEEP);
  if (rule === RULE.TRASH) return r_('Sender rule: Trash', ACTION.TRASH);
  if (rule === RULE.REVIEW) return labelOrKeep_(m, 'Sender rule: Review', L.review);

  if (matches_(s.re.jobSender, m.email) || matches_(s.re.jobSubject, m.subject)) {
    return m.date < daysAgo_(s[SETTING.JOB_DAYS]) ? r_('Job search (old)', ACTION.TRASH) : r_('Job search', ACTION.KEEP);
  }

  if (REPLY_PREFIX_RE.test(m.subject)) return labelOrKeep_(m, 'Reply from a person', L.review);

  if (m.unsub) {
    const decision = (ctx.news[m.email] || {}).decision || DECISION.PENDING;
    if (decision === DECISION.UNSUBSCRIBE || decision === DECISION.INACTIVE) {
      return r_('Newsletter (' + decision.toLowerCase() + ')', ACTION.TRASH);
    }
    // Kept and undecided newsletters: only recent issues are kept, nothing is unsubscribed
    const kind = decision === DECISION.KEEP ? 'kept' : 'waiting for decision';
    return m.date < daysAgo_(s[SETTING.NEWSLETTER_DAYS])
      ? r_('Newsletter (' + kind + ', old issue)', ACTION.TRASH) : r_('Newsletter (' + kind + ')', ACTION.KEEP);
  }

  const isAuto = matches_(s.re.auto, m.email.split('@')[0]);
  if (!isAuto && matches_(s.re.personal, m.email)) return r_('Personal sender', ACTION.KEEP);
  if (isAuto) return r_('Automated notification', ACTION.TRASH);
  return labelOrKeep_(m, 'Unknown sender', L.review);
}

function r_(rule, action, label) {
  return { rule, action, label: label || '' };
}

function labelOrKeep_(m, rule, label) {
  return m.labels.includes(label) ? r_(rule, ACTION.KEEP) : r_(rule, ACTION.LABEL, label);
}

function isUserLabel_(name, ownLabels) {
  return SYSTEM_LABELS.indexOf(name) < 0 && name.indexOf('CATEGORY_') !== 0 && ownLabels.indexOf(name) < 0;
}

/** Senders listed in Private.gs are never trashed. Checked in classification AND before every trash. */
function isProtectedSender_(email, name) {
  const p = privateConfig_();
  return (!!p.PROTECTED_ADDRESSES && p.PROTECTED_ADDRESSES.test(String(email || ''))) ||
    (!!p.PROTECTED_NAMES && p.PROTECTED_NAMES.test(String(name || '')));
}

function senderRule_(email, rules) {
  for (const r of rules) {
    if (r.pattern.indexOf('@') > 0 ? email === r.pattern : domainMatch_(email, r.pattern.replace(/^@/, ''))) {
      return r.action;
    }
  }
  return '';
}

function domainMatch_(email, domain) {
  const d = email.split('@')[1] || '';
  return d === domain || d.endsWith('.' + domain);
}

// ---------- Gmail actions ----------

const TRASHED = 'Trashed';

/** Threads you have written in, and the IDs of this tool's own labels – fetched fresh before acting. */
function loadProtection_(s) {
  const sentThreads = new Set(listIds_('in:sent').map(m => m.threadId));
  const ownLabelIds = new Set((Gmail.Users.Labels.list('me').labels || [])
    .filter(l => s.ownLabels.indexOf(l.name) >= 0).map(l => l.id));
  return { sentThreads, ownLabelIds, me: Gmail.Users.getProfile('me').emailAddress.toLowerCase() };
}

/** Moves a message to trash after a last check of its current sender and labels. */
function safeTrash_(id, threadId, prot) {
  if (prot.sentThreads.has(threadId)) return 'Skipped: thread you wrote in';
  let msg;
  try {
    msg = withRetry_(() => Gmail.Users.Messages.get('me', id, { format: 'metadata', metadataHeaders: ['From'] }));
  } catch (e) {
    return /not found/i.test(e.message) ? 'No longer exists' : 'Error: ' + e.message.slice(0, 100);
  }
  const from = parseFrom_(headerMap_(msg)['from']);
  if (isProtectedSender_(from.email, from.name)) return 'Skipped: protected sender';
  const labels = msg.labelIds || [];
  if (labels.includes('TRASH')) return 'Already in trash';
  if (labels.includes('STARRED')) return 'Skipped: starred';
  if (labels.some(l => l.indexOf('Label_') === 0 && !prot.ownLabelIds.has(l))) return 'Skipped: has your own label';
  withRetry_(() => Gmail.Users.Messages.trash('me', id));
  return TRASHED;
}

/** Adds labels in batches; if Gmail rejects a batch, the messages are tried one by one. */
function applyLabels_(toLabel, cache) {
  Object.keys(toLabel).forEach(name => {
    const id = labelId_(name, cache);
    for (let i = 0; i < toLabel[name].length; i += 1000) {
      const ids = toLabel[name].slice(i, i + 1000);
      try {
        withRetry_(() => Gmail.Users.Messages.batchModify({ ids, addLabelIds: [id] }, 'me'));
      } catch (e) {
        ids.forEach(msgId => {
          try { withRetry_(() => Gmail.Users.Messages.modify({ addLabelIds: [id] }, 'me', msgId)); } catch (e2) { /* message gone */ }
        });
      }
    }
  });
}

function isValidId_(id) {
  return typeof id === 'string' && /^[0-9a-f]+$/i.test(id);
}

function labelId_(name, cache) {
  if (cache[name]) return cache[name];
  if (name.indexOf('/') > 0) labelId_(name.slice(0, name.lastIndexOf('/')), cache);   // parent first
  const found = (Gmail.Users.Labels.list('me').labels || []).find(l => l.name === name);
  const id = found ? found.id
    : Gmail.Users.Labels.create({ name, labelListVisibility: 'labelShow', messageListVisibility: 'show' }, 'me').id;
  cache[name] = id;
  return id;
}

/**
 * Unsubscribes newsletters marked "Unsubscribe" that have no status yet. In a weekly run in
 * "Execute" mode the sender's older messages are trashed too. Returns false if time ran out.
 */
function processNewsletterDecisions_(t0, trashOld) {
  const sh = sheet_(SHEETS.NEWS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return true;
  const data = sh.getRange(2, 1, n, NEWS_COLS.length).getValues();
  const today = Utilities.formatDate(new Date(), CFG.TIMEZONE, 'yyyy-MM-dd');
  let ctx = null;
  for (let i = 0; i < n; i++) {
    const [email, , , , , decision, status] = data[i];
    if (status && String(status) !== NEW_STATUS) continue;
    if (decision !== DECISION.UNSUBSCRIBE && decision !== DECISION.INACTIVE) continue;
    if (Date.now() - t0 > CFG.MAX_RUN_MS) return false;

    const addr = String(email).toLowerCase();
    let result = decision === DECISION.UNSUBSCRIBE ? unsubscribe_(addr) : 'Inactive – not unsubscribed';
    if (trashOld) {
      ctx = ctx || context_();
      const days = Number(ctx.s[SETTING.PROTECT_DAYS]);
      const ids = listIds_('from:' + addr + ' older_than:' + days + 'd');
      const moved = ids.filter(m => safeTrash_(m.id, m.threadId, ctx.prot) === TRASHED).length;
      result += '; ' + moved + ' messages to trash';
    }
    sh.getRange(i + 2, 7).setValue(today + ': ' + result);
  }
  return true;
}

/** Unsubscribes via one-click (RFC 8058) or an unsubscribe email. Plain links are never opened automatically. */
function unsubscribe_(email) {
  const page = withRetry_(() => Gmail.Users.Messages.list('me', { q: 'from:' + email, maxResults: 1, includeSpamTrash: true }));
  const first = (page.messages || [])[0];
  if (!first) return 'No message found';
  const msg = withRetry_(() => Gmail.Users.Messages.get('me', first.id,
    { format: 'metadata', metadataHeaders: ['List-Unsubscribe', 'List-Unsubscribe-Post'] }));
  const h = headerMap_(msg);

  const urls = (String(h['list-unsubscribe'] || '').match(/<[^>]+>/g) || []).map(u => u.slice(1, -1).trim());
  const http = urls.find(u => /^https:/i.test(u)) || urls.find(u => /^http:/i.test(u));
  const mailto = urls.find(u => /^mailto:/i.test(u));

  if (http && /one-click/i.test(h['list-unsubscribe-post'] || '')) {
    try {
      const res = UrlFetchApp.fetch(http, { method: 'post', payload: 'List-Unsubscribe=One-Click', muteHttpExceptions: true });
      if (res.getResponseCode() < 400) return 'Unsubscribed (one-click)';
    } catch (e) { /* fall back to mail */ }
  }
  if (mailto) {
    const parts = mailto.slice(7).split('?');
    const params = {};
    (parts[1] || '').split('&').forEach(p => {
      const kv = p.split('=');
      if (kv[0]) params[kv[0].toLowerCase()] = decodeURIComponent((kv[1] || '').replace(/\+/g, ' '));
    });
    sendMail_(decodeURIComponent(parts[0]), params.subject || 'unsubscribe', params.body || 'unsubscribe');
    return 'Unsubscribed (email sent)';
  }
  if (http) return 'Unsubscribe manually: ' + http;
  return 'No unsubscribe header';
}

function sendMail_(to, subject, body) {
  const clean = s => String(s).replace(/[\r\n]+/g, ' ').trim();   // no header injection
  const b64 = s => Utilities.base64Encode(s, Utilities.Charset.UTF_8);
  const raw = 'To: ' + clean(to) + '\r\n' +
    'Subject: =?UTF-8?B?' + b64(clean(subject)) + '?=\r\n' +
    'Content-Type: text/plain; charset=UTF-8\r\n' +
    'Content-Transfer-Encoding: base64\r\n\r\n' + b64(body);
  withRetry_(() => Gmail.Users.Messages.send({ raw: Utilities.base64EncodeWebSafe(raw) }, 'me'));
}

// ---------- Reading ----------

function readRaw_() {
  const sh = sheet_(SHEETS.RAW);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const seen = new Set();
  return sh.getRange(2, 1, n, RAW_COLS.length).getValues()
    .filter(r => r[0] && !seen.has(r[0]) && seen.add(r[0]))
    .map(rowToMsg_);
}

function rowToMsg_(r) {
  return {
    id: String(r[0]),
    threadId: String(r[1]),
    date: r[2] instanceof Date ? r[2] : new Date(r[2]),
    email: String(r[3]).replace(/^'/, '').toLowerCase(),
    name: String(r[4]).replace(/^'/, ''),
    subject: String(r[6]).replace(/^'/, ''),
    labels: String(r[8]).split(', ').filter(String),
    unsub: String(r[9])
  };
}

function fetchMsgs_(q, labelNames) {
  return fetchMetadata_(listIds_(q).map(m => m.id)).map(m => rowToMsg_(toRow_(m, labelNames)));
}

function readSettings_() {
  const sh = sheet_(SHEETS.SETTINGS);
  const s = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { s[String(r[0]).trim()] = r[1]; });
  }
  DEFAULT_SETTINGS.forEach(d => { if (s[d[0]] === undefined) s[d[0]] = d[1]; });
  s.re = {
    receipt: regex_(s[SETTING.RECEIPT_WORDS]),
    cases: regex_(s[SETTING.CASE_WORDS]),
    jobSender: regex_(s[SETTING.JOB_SENDERS]),
    jobSubject: regex_(s[SETTING.JOB_SUBJECTS]),
    personal: regex_(s[SETTING.PERSONAL_DOMAINS]),
    auto: regex_(s[SETTING.AUTO_SENDERS])
  };
  s.labels = {
    receipts: String(s[SETTING.LABEL_RECEIPTS]),
    cases: String(s[SETTING.LABEL_CASES]),
    review: String(s[SETTING.LABEL_REVIEW])
  };
  // The tool's own labels and their parents do not count as "your own label"
  s.ownLabels = [];
  Object.values(s.labels).forEach(l => {
    const parts = l.split('/');
    for (let i = 1; i <= parts.length; i++) s.ownLabels.push(parts.slice(0, i).join('/'));
  });
  return s;
}

function readRules_() {
  const sh = sheet_(SHEETS.RULES);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const actions = Object.values(RULE);
  return sh.getRange(2, 1, n, 2).getValues()
    .map(r => ({ pattern: String(r[0]).trim().toLowerCase(), action: String(r[1]).trim() }))
    .filter(r => r.pattern && actions.indexOf(r.action) >= 0);
}

function readNewsletters_() {
  const sh = sheet_(SHEETS.NEWS);
  const out = {};
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    sh.getRange(2, 1, n, NEWS_COLS.length).getValues().forEach(r => {
      if (r[0]) out[String(r[0]).toLowerCase()] = { decision: String(r[5] || DECISION.PENDING) };
    });
  }
  return out;
}

// ---------- Log and status ----------

function logRow_(id, email, subject, rule, result) {
  return [new Date(), String(id), safe_(email), safe_(String(subject).slice(0, 150)), rule, result];
}

function setCleanupStatus_(text) {
  sheet_(SHEETS.STATUS).getRange(7, 1, 2, 2).setValues([['Clean-up', text], ['Clean-up updated', new Date()]]);
}
