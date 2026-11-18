<img src="docs/logo.svg" alt="arnoldus.tech" width="220">

# Gmail Cleanup for Google Workspace

A rule-based clean-up of a personal Gmail account that keeps what you may need later – correspondence, receipts, cases – and clears out notifications, alerts and newsletters you never asked to keep. It runs entirely in your own Google account: a **Google Sheet** holds the audit, the rules and the log, and **Google Apps Script** does the work, once by hand and then every week by itself.

> Status: in use on a personal mailbox. No mail data is included in this repository.

## Why

Most of a long-lived mailbox is noise: delivery notices, login alerts, calendar pings, job alerts, newsletters from shops visited once. Deleting by hand doesn't scale, and generic "inbox zero" tools tend to delete too much or ask for access to your mail on someone else's server.

This project takes the cautious route. It reads headers only, shows you a plan before it touches anything, never deletes permanently, and treats anything written by a person as worth keeping.

## What it does

| | |
|---|---|
| **Audit (read-only)** | Reads the headers of every message and builds an overview per sender and per domain: count, read rate, starred, replied-in-thread, Gmail category, unsubscribe method, size and date range |
| **Rules** | Global rules (protection period, receipts, cases, job search, newsletters, automated senders) plus a per-sender list – *Keep*, *Trash* or *Review* – edited in the sheet |
| **Dry-run plan** | Classifies every message and writes a *Plan* and a *Summary* (counts per rule, top senders to be trashed, number of receipts) before anything is changed |
| **Execution** | Moves to trash, adds labels and unsubscribes in resumable chunks, with a fresh safety check of each message right before it is trashed. Every action is logged |
| **Newsletter allow-list** | Each newsletter sender is *Keep*, *Unsubscribe* or *Inactive*. New newsletters are registered as **?** and never unsubscribed automatically – you get a weekly email asking you to decide. Issues older than 90 days are trashed either way |
| **Unsubscribe** | RFC 8058 one-click (`List-Unsubscribe-Post`) or an unsubscribe email. Plain unsubscribe links are never opened automatically – they are listed for you to click |
| **Weekly run** | Every Monday: registers new newsletters, carries out your decisions and cleans up messages that have just left the protection period. Starts in *Log only* mode |

## What is never trashed

In this order – the first match wins:

1. **Protected senders** listed in `Private.gs` (e.g. family), whatever the rules say – checked during classification *and* again right before every trash
2. Your own messages, and anything inside a **thread you wrote in**
3. Messages **newer than the protection period** (30 days)
4. **Starred** messages and messages with **any label of your own**
5. **Receipts** (invoice, order, booking, contract, policy …) – kept forever and labelled
6. Subjects matching your **case words** – kept and labelled
7. Senders with a **Keep** rule, **personal senders** on free-mail domains, **replies** (`Re:`, `SV:`, `Fwd:` …) and **job search** mail (for a year)
8. Anything the rules don't recognise is **labelled for review**, not trashed

## Architecture

```
Google Sheet (bound Apps Script)
 ├─ Raw data        one row per message: IDs, date, sender, subject, labels, unsubscribe method
 ├─ Senders         overview per sender
 ├─ Domains         overview per domain
 ├─ Settings        retention periods, weekly mode, labels and the regexes used by the rules
 ├─ Sender rules    Keep / Trash / Review per address or domain
 ├─ Newsletters     Keep / Unsubscribe / Inactive / ? per newsletter sender
 ├─ Plan            what the dry run decided per message
 ├─ Summary         counts per rule and top senders
 ├─ Log             every action taken (or that would be taken in Log only mode)
 └─ Status          audit and clean-up progress

Gmail labels  Archive/Receipts · Archive/Cases · Review   (names configurable)
```

| Source file | Responsibility |
|---|---|
| `Config.gs` | Constants, sheet layouts, default settings and sender rules |
| `Utils.gs` | Sheet helpers, paced Gmail reads, header parsing, triggers |
| `Audit.gs` | Read-only audit, sender and domain overview, compact export |
| `Cleanup.gs` | Classification, dry-run plan, execution, unsubscribe, weekly run |
| `Menu.gs` | The spreadsheet menu |
| `Private.gs` | Your protected senders – **git-ignored**, see `Private.gs.example` |

## Design principles

- **Headers only.** The audit and the rules use the `From`, `Subject` and `List-Unsubscribe` headers. Message bodies are never read.
- **Plan before action.** The first clean-up is always a dry run you can inspect in the sheet. The weekly run starts in *Log only*.
- **Nothing permanent.** "Trash" is Gmail's trash (restorable for 30 days). The tool has no code path that deletes permanently.
- **People over patterns.** A thread you replied in, a reply from someone else, a personal address or a protected sender always beats a rule.
- **Your account, your data.** Everything runs as you, in your Google account. Nothing is sent anywhere except unsubscribe requests to the senders you chose to unsubscribe from.
- **Quota-aware.** Reads are paced below Gmail's per-minute quota and use the Gmail advanced service rather than UrlFetch, whose daily quota a mailbox audit quickly exhausts. Long jobs resume themselves through triggers.

## Setup

Prerequisites: a Google account, [Node.js](https://nodejs.org/) and [clasp](https://github.com/google/clasp).

1. Create an empty Google Sheet (e.g. `Gmail cleanup`).
2. Enable the Apps Script API at <https://script.google.com/home/usersettings>.
3. Clone this repo, add your protected senders and bind it to the sheet:
   ```bash
   cp src/Private.gs.example src/Private.gs   # then edit it
   npm install -g @google/clasp
   clasp login
   clasp create-script --parentId <spreadsheet-id> --rootDir src   # or copy .clasp.json.example to .clasp.json
   clasp push
   ```
4. Reload the sheet. In **🧹 Gmail cleanup**:
   1. **Audit: start** – reads the mailbox in chunks (≈20–40 minutes for 20,000 messages). Follow it in *Status*.
   2. **1. Prepare rules** – creates *Settings*, *Sender rules* and *Newsletters*. Mark the newsletters you want with **Keep**.
   3. **2. Build plan (dry run)** – review *Summary* and *Plan*. Adjust rules and rebuild until it looks right.
   4. **3. Execute plan** – runs in the background; you can close the sheet.
   5. **Activate weekly run**, check the *Log* after the first Monday, then set *Weekly mode* to **Execute**.

Without clasp you can also create the files by hand in **Extensions → Apps Script** and paste the contents of `src/` (show `appsscript.json` under *Project settings*).

User guide: [docs/USER_GUIDE.md](docs/USER_GUIDE.md)

## Troubleshooting

- **"Service invoked too many times for one day: urlfetch"** – an older parallel version used UrlFetch for reads. This version uses the Gmail advanced service; just resume.
- **"Quota exceeded … per minute per user"** – lower `CFG.MAX_PER_MINUTE` in `Config.gs` and resume.
- **"Invalid ids value"** – Sheets turned an all-digit message ID into a number. Such rows are skipped and logged; new plans store IDs as text.

## Privacy note

All mail data stays in your own Google account. The repository contains code only – no mail, no sheet data, no IDs and no protected senders. `Private.gs`, `.clasp.json` (script binding), `.clasprc.json` (credentials) and any CSV exports are git-ignored.

## Licence

[MIT](LICENSE)
