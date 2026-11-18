# User guide

## The weekly routine (≈5 minutes)

1. When the email **"Gmail cleanup: N new newsletters"** arrives, open the sheet's **Newsletters** sheet and set each **?** to **Keep** or **Unsubscribe**. Until you decide, nothing is unsubscribed – only issues older than 90 days are trashed.
2. Now and then, look through the Gmail label **Review**. Delete what you don't need, and add a row to **Sender rules** for senders that keep coming back.
3. Once a month, check **Status** to see that the weekly run is still working.

## Sheets

| Sheet | What to do there |
|---|---|
| **Settings** | Retention periods, *Weekly mode* (`Log only` / `Execute`), the weekly email, label names and the regexes behind the rules |
| **Sender rules** | One row per address or domain: **Keep**, **Trash** or **Review**. A domain also matches its subdomains. First match wins |
| **Newsletters** | One row per newsletter sender: **Keep**, **Unsubscribe**, **Inactive** (no longer sends; no unsubscribe needed) or **?** (new). *Status* shows when and how it was unsubscribed |
| **Plan** / **Summary** | The result of the last dry run |
| **Log** | Every action. In *Log only* mode the results read "(would trash)" / "(would label)" |
| **Status** | Progress of the audit and the clean-up, and the last error if any |

## Typical adjustments

- **Something useful was going to be trashed** – add the sender to *Sender rules* with **Keep**, or add a word to *Receipt words* or *Case words* in *Settings*. Rebuild the plan to check.
- **Too much ends up in Review** – add the senders to *Sender rules* with **Trash**.
- **A person must never be touched** – add their addresses and name to `Private.gs` and push again. This check overrides everything else.

## Restoring

Everything the tool trashes is in Gmail's **Trash** for 30 days. The *Log* sheet shows the message ID, sender and subject of each action, so you can find a message and move it back.

## Unsubscribing

- **one-click** – the tool sends the standard one-click request to the sender.
- **mail** – the tool sends a short unsubscribe email from your account to the sender's unsubscribe address.
- **link** – the tool never opens such links; the *Status* column in *Newsletters* shows "Unsubscribe manually" with the link for you to click.
