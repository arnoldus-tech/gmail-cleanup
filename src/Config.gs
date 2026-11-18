/**
 * Gmail Cleanup – configuration and defaults.
 *
 * A rule-based clean-up of a personal Gmail account, bound to a Google Sheet
 * that holds the audit, the rules and the log. Nothing is ever deleted
 * permanently: "Trash" means Gmail's trash, which can be restored for 30 days.
 *
 * Personal settings (senders that must never be touched) live in Private.gs,
 * which is git-ignored. See Private.gs.example.
 */

const CFG = {
  TIMEZONE: 'Europe/Copenhagen',
  PAGE_SIZE: 500,                        // messages per Gmail list page
  MAX_PER_MINUTE: 2000,                  // stay below Gmail's per-user quota per minute
  MAX_RUN_MS: 3.5 * 60 * 1000,           // Apps Script stops executions after 6 minutes
  MAX_FAILED_RUNS: 5,                    // consecutive failures before a resumable job gives up
  COMPACT_MIN_COUNT: 3,                  // senders with fewer messages are left out of the compact export
  WEEKLY_HOUR: 6                         // weekly run: Monday between 06 and 07
};

const SHEETS = {
  RAW: 'Raw data', SENDERS: 'Senders', DOMAINS: 'Domains', STATUS: 'Status', ERRORS: 'Errors',
  COMPACT: 'Compact senders', SETTINGS: 'Settings', RULES: 'Sender rules', NEWS: 'Newsletters',
  PLAN: 'Plan', SUMMARY: 'Summary', LOG: 'Log'
};

/** Headers fetched per message. The body is never read. */
const META_HEADERS = ['From', 'Subject', 'List-Unsubscribe', 'List-Unsubscribe-Post'];

const RAW_COLS = ['ID', 'Thread ID', 'Date', 'Email', 'Name', 'Domain', 'Subject', 'Bytes', 'Labels', 'Unsubscribe'];
const SENDER_COLS = ['Sender', 'Name', 'Domain', 'Count', 'In inbox', 'Unread', 'Read %', 'Starred', 'Important',
  'Replied in thread', 'Category', 'Unsubscribe', 'MB', 'Oldest', 'Newest', 'Example subjects'];
const DOMAIN_COLS = ['Domain', 'Senders', 'Count', 'Read %', 'Replied in thread', 'MB', 'Newest'];
const COMPACT_COLS = ['Sender', 'Name', 'Count', 'Read %', 'Starred', 'Replied', 'Category', 'Unsubscribe', 'MB',
  'From', 'To', 'Subjects'];
const PLAN_COLS = ['ID', 'Thread ID', 'Date', 'Sender', 'Subject', 'Rule', 'Action'];
const NEWS_COLS = ['Sender', 'Name', 'Count', 'Latest', 'Method', 'Decision', 'Status'];
const LOG_COLS = ['Time', 'ID', 'Sender', 'Subject', 'Rule', 'Result'];
const RULE_COLS = ['Pattern', 'Action', 'Note'];
const SETTING_COLS = ['Setting', 'Value', 'Explanation'];

const CATEGORY_NAMES = {
  CATEGORY_PERSONAL: 'Primary', CATEGORY_PROMOTIONS: 'Promotions', CATEGORY_SOCIAL: 'Social',
  CATEGORY_UPDATES: 'Updates', CATEGORY_FORUMS: 'Forums'
};
const UNSUB_RANK = { '': 0, mail: 1, link: 2, 'one-click': 3 };
const SYSTEM_LABELS = ['INBOX', 'UNREAD', 'IMPORTANT', 'SENT', 'DRAFT', 'SPAM', 'TRASH', 'STARRED', 'CHAT'];

/** What the classifier decides for one message. */
const ACTION = { KEEP: 'Keep', TRASH: 'Trash', LABEL: 'Label' };
/** Actions a row in "Sender rules" can have. */
const RULE = { KEEP: 'Keep', TRASH: 'Trash', REVIEW: 'Review' };
/** Decisions in "Newsletters". "?" = new, not decided yet: never unsubscribed automatically. */
const DECISION = { KEEP: 'Keep', UNSUBSCRIBE: 'Unsubscribe', INACTIVE: 'Inactive', PENDING: '?' };
const NEW_STATUS = 'New – waiting for you';

const SETTING = {
  PROTECT_DAYS: 'Protection period (days)',
  NEWSLETTER_DAYS: 'Newsletter retention (days)',
  JOB_DAYS: 'Job search retention (days)',
  WEEKLY_MODE: 'Weekly mode',
  WEEKLY_MAIL: 'Weekly email to me (YES/NO)',
  LABEL_RECEIPTS: 'Label for receipts',
  LABEL_CASES: 'Label for cases',
  LABEL_REVIEW: 'Label for review',
  RECEIPT_WORDS: 'Receipt words (regex)',
  CASE_WORDS: 'Case words (regex)',
  JOB_SENDERS: 'Job search senders (regex)',
  JOB_SUBJECTS: 'Job search subjects (regex)',
  PERSONAL_DOMAINS: 'Personal mail domains (regex)',
  AUTO_SENDERS: 'Automated senders (regex)'
};
const MODE = { LOG_ONLY: 'Log only', EXECUTE: 'Execute' };

const DEFAULT_SETTINGS = [
  [SETTING.PROTECT_DAYS, 30, 'Messages newer than this are never touched'],
  [SETTING.NEWSLETTER_DAYS, 90, 'Newsletters you keep or have not decided on: issues older than this are trashed'],
  [SETTING.JOB_DAYS, 365, 'Messages about job applications are trashed after this'],
  [SETTING.WEEKLY_MODE, MODE.LOG_ONLY, '"Log only" (writes to Log what would happen) or "Execute"'],
  [SETTING.WEEKLY_MAIL, 'YES', 'Email yourself when new newsletters are waiting for a decision'],
  [SETTING.LABEL_RECEIPTS, 'Archive/Receipts', 'Receipts are kept forever with this label'],
  [SETTING.LABEL_CASES, 'Archive/Cases', 'Messages matching the case words are kept with this label'],
  [SETTING.LABEL_REVIEW, 'Review', 'Unknown senders get this label instead of being trashed'],
  [SETTING.RECEIPT_WORDS,
    'receipt|invoice|order|booking|reservation|itinerary|e-ticket|boarding|subscription (is )?confirmed|contract|policy|warranty|credit note|refund|payslip|' +
    'kvittering|faktura|ordre|bekræftelse på (bestilling|booking|køb)|opgørelse|police|forsikringsbevis|kontrakt|købsaftale|kendelse|lønseddel|kreditnota|refusion|garanti',
    'Subjects matching this are always kept and get the receipts label'],
  [SETTING.CASE_WORDS, '', 'Optional, e.g. a property address or a case number. Matching subjects are kept with the cases label'],
  [SETTING.JOB_SENDERS,
    'teamtailor|successfactors|myworkday|jobs2web|greenhouse|lever\\.co|smartrecruiters|workable|recruitee|talentech|emply|hr-manager\\.net|jobs-noreply@linkedin|hit-reply@linkedin',
    'Applicant tracking systems and recruiter messages'],
  [SETTING.JOB_SUBJECTS,
    'your application|job application|application for|thank you for your interest|interview|candidate|resume|\\bcv\\b|keep your profile|' +
    'din ansøgning|ansökan|tak for din interesse|jobsamtale|kandidat|emnebank',
    'Subjects about job applications'],
  [SETTING.PERSONAL_DOMAINS,
    '@(gmail|googlemail|hotmail|outlook|live|yahoo|icloud|me|mail|msn|proton|protonmail)\\.(com|dk|me|net)$',
    'People on these domains are kept'],
  [SETTING.AUTO_SENDERS,
    '^(no-?reply|no_reply|donotreply|do-not-reply|do_not_reply|notification|automail|automated|mailer|info|news|hello|hi|support|service|customer|webshop|shop|store|system|account|express|post|e-mail|email|security|alert|calendar|marketing|survey|tracking|bounce|receipts|reservation|login|profile|online|sales|contact|mail|' +
    'kanikkebesvares|kan-ikke-besvares|besvaresikke|nyhed|hej|kundeservice|kundesupport|besked|salg|kontakt)',
    'Start of the sender address for automated mail; trashed after the protection period']
];

/**
 * Starting rules for a new sheet: [pattern, action, note]. First match wins.
 * Pattern = a full address, or a domain (which also matches its subdomains).
 */
const DEFAULT_SENDER_RULES = [
  ['calendar-notification@google.com', RULE.TRASH, 'Calendar notifications'],
  ['accounts.google.com', RULE.TRASH, 'Security notices'],
  ['noreply-local-guides@google.com', RULE.TRASH, ''],
  ['google-maps-noreply@google.com', RULE.TRASH, ''],
  ['noreply-maps-timeline@google.com', RULE.TRASH, ''],
  ['security-noreply@linkedin.com', RULE.TRASH, 'Security notices'],
  ['messaging-digest-noreply@linkedin.com', RULE.TRASH, 'Message notifications'],
  ['jobalerts-noreply@linkedin.com', RULE.TRASH, 'Job alerts'],
  ['invitations@linkedin.com', RULE.TRASH, ''],
  ['trustpilotmail.com', RULE.TRASH, 'Review requests']
];

/** "SV:", "Re:", "Fwd:" … – written by a person, even when the address looks automated. */
const REPLY_PREFIX_RE = /^\s*(re|sv|aw|vs|fw|fwd|wg)\s*:/i;

/** Senders that are never trashed, from Private.gs (git-ignored). Safe if the file is missing. */
function privateConfig_() {
  return typeof PRIVATE === 'undefined' ? {} : PRIVATE;
}
