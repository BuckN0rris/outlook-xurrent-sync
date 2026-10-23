# outlook-xurrent-sync

Keeps the categories of your Outlook mails in step with the status of the
[Xurrent](https://www.xurrent.com/) requests they belong to.

English · [Deutsch](README.de.md) · unofficial, not affiliated with Microsoft or Xurrent

Xurrent mails you when a request is assigned to you and when someone adds a note. You
categorise those mails ("To do"), work the ticket, and then have to categorise the mails
again by hand. This script does the second half: it looks at every mail that still
carries an open category, finds the request number in it, asks Xurrent for the status,
and recategorises the mail. It runs locally, on demand or on a schedule. No server, no
service, no cloud component.

| Request in Xurrent | Mail in Outlook | What happens |
|---|---|---|
| completed | has an open category | open categories removed, done category added |
| waiting for a customer or someone else | has the open category | moved to the "waiting" category (optional) |
| anything else, unreadable, or no request number found | any | left alone, listed in the report |

## Guarantees

| | |
|---|---|
| **Dry run by default** | Without `--apply` nothing is written to Outlook. Local state and logs are still kept. |
| **Xurrent is read-only** | `XurrentClient._call` can only send `GET`. No create, update or comment. |
| **Outlook: only `categories`** | The one write is `PATCH /me/messages/{id}` with the `categories` field and nothing else. No move, delete, flag, reply or forward. |
| **Other categories survive** | Categories you set yourself stay, in their order. Only the configured "open" categories are removed. |
| **No request number, no change** | Mails without a recognised number are never touched; they are listed for a manual look. |
| **A note after closure stays open** | A mail that arrived clearly *after* the request was closed is never marked done, it may still need an answer. |
| **Idempotent** | A second run right after the first changes nothing. |
| **Your own mailbox only** | Delegated access as the signed-in user. No application permissions, no access to other mailboxes. |

## Requirements

- Python 3.10 or newer
- A Microsoft 365 / Exchange Online mailbox (Microsoft Graph)
- A Xurrent account and a Personal Access Token with `Request:Read`

Built and used on Windows 11. The sync itself is plain Python; the double-click launchers
(`sync-*.cmd`) and the scheduling instructions below are Windows-specific.

## Quick start

```bash
python -m pip install -r requirements.txt
```

Copy `.env.example` to `.env` and fill it in. Six lines are enough, everything else has a
default (see [Configuration](#configuration)).

```bash
python sync.py --check
```

Checks the Xurrent token, signs you in to Microsoft (a browser window opens the first
time) and verifies that your category names exist. Writes nothing.

```bash
python sync.py
```

Dry run over every mail that still has an open category. Prints what it would change:

```
REQUEST     STATUS                ACTION       SUBJECT
----------  --------------------  -----------  ----------------------------
80000001    completed             WOULD-APPLY  WG: Login details for mailbox
80000003    waiting_for_customer  WOULD-APPLY  New employee - onboarding
80000002    in_progress           STILL-OPEN   Printer on 2nd floor offline
70000001    completed             AFTER-CLOSE  Note after closure
-           -                     NO-ID        Could you look at the printer...
```

Then apply carefully, then regularly:

```bash
python sync.py --apply --limit 5
```

### Xurrent token

Avatar → **My Profile** → **Personal Access Tokens** → *Generate new token*. In the scope
table use effect `Allow`, record type **Request**, action **Read**, and (recommended) a
condition **Account** = your account id. A token never has more rights than the person
who created it.

### Microsoft Graph sign-in

By default the script signs in with Microsoft's public "Microsoft Graph Command Line
Tools" client (`GRAPH_CLIENT_ID`). That works as long as your tenant lets users consent to
the delegated scopes `Mail.ReadWrite` (needed to write `categories`; `Mail.Read` is not
enough) and `MailboxSettings.Read` (only to check that your category names exist; with
`VALIDATE_CATEGORIES=false` it is not needed).

If your tenant blocks that, register your own app: *Mobile and desktop application*,
redirect URI `http://localhost`, "Allow public client flows" = yes, **delegated**
permissions only, then set `GRAPH_CLIENT_ID` and `GRAPH_TENANT_ID`.

## How a mail is judged

- **Several mails per request.** Xurrent mails on assignment and on every note, so one
  request often has several mails. When it is completed, all of them are recategorised.
- **Notes on a closed request.** A mail counts as done only if the request's
  `completed_at` is not more than `AFTER_CLOSE_GRACE_SECONDS` (default two hours) before
  the mail's `receivedDateTime`. The grace window exists because Xurrent sends the
  "completed" notification *after* the closure; without it that very mail would be the one
  left open. Everything later shows up as `AFTER-CLOSE` and is not touched.
- **Waiting requests.** With `WAITING_CATEGORY` set, mails of requests in `waiting_for` or
  `waiting_for_customer` move from the open category to the waiting category, so "To do"
  holds only what is waiting for *you*. All requests of the mail must be waiting or done.
  Mentions (`You were mentioned in Request #…`) stay where they are, because someone asked
  you something. The waiting category stays an open category: once the request completes,
  the next run sets the mail to done. The move goes one way only, a request returning to
  `in_progress` does not move the mail back.
- **Several requests in one mail.** `MULTI_REQUEST_POLICY=all` (default) needs every
  request done, `any` needs one.

`python sync.py --explain "text"` shows, for each mail whose subject contains `text`, what
was found and why it was or was not touched. It never writes.

## How the request number is found

Four steps, first hit wins ([extract.py](outlook_xurrent/extract.py)):

| # | Source | Pattern |
|---|---|---|
| 1 | Subject | `#\s*(\d{4,})` |
| 2 | Link in the raw HTML body | `<host>/(?:self-service/)?requests/(\d+)` |
| 3 | Body text with HTML removed | `#\s*(\d{4,})` |
| 4 | `X-Xurrent-*`, `Message-ID`, `References` headers | `requests?[._/-](\d{4,})` |

Every pattern needs a `#` or `/requests/` in front of the digits, so a customer number
sitting next to the request number is never mistaken for it. `XURRENT_INSTANCE_HOST` is an
allow-list for the links in step 2, so a forwarded link to someone else's Xurrent account
is ignored. Bodies and headers are fetched only for candidate mails, once per mail; the
result is cached in `.state/state.json`.

## Configuration

Settings come from the environment or from `.env`; real environment variables win. Timeouts,
retries and page sizes are constants in [config.py](outlook_xurrent/config.py) on purpose.

| Variable | Default | Meaning |
|---|---|---|
| `XURRENT_TOKEN` | – | **Required.** Personal Access Token, scope `Request:Read`. |
| `XURRENT_ACCOUNT` | – | **Required.** Value of the `X-Xurrent-Account` header, usually the instance subdomain. |
| `DONE_CATEGORY` | – | **Required.** Exact name of the "done" category. |
| `OPEN_CATEGORY` | empty | Exact names of the "open" categories, comma separated. All are removed when a mail becomes done. Empty: every mail without `DONE_CATEGORY` counts as open. |
| `WAITING_CATEGORY` | empty | Category for mails whose requests wait for someone else. Must also be in `OPEN_CATEGORY`. Empty: feature off. |
| `WAITING_STATUSES` | `waiting_for,waiting_for_customer` | Statuses that count as waiting. |
| `DONE_STATUSES` | `completed` | Statuses that count as done. Valid: `declined`, `on_backlog`, `assigned`, `accepted`, `in_progress`, `waiting_for`, `waiting_for_customer`, `reservation_pending`, `workflow_pending`, `project_pending`, `completed`. |
| `MULTI_REQUEST_POLICY` | `all` | `all` or `any`, see above. |
| `AFTER_CLOSE_GRACE_SECONDS` | `7200` | How long after `completed_at` a mail still belongs to the closure. `0`: strict. |
| `XURRENT_INSTANCE_HOST` | empty | Allow-list of hosts for link detection, comma separated (agent instance and customer portal). Empty: any `*.xurrent.com`. |
| `XURRENT_API_ROOT` | `https://api.xurrent.com/v1` | Regional API root (`api.au`, `api.uk`, `api.ch`, `api.us`). |
| `LOOKBACK_DAYS` | `0` | Only mails from the last N days. `0`: no limit. `--days` overrides it. |
| `VALIDATE_CATEGORIES` | `true` | Check at start that the category names exist in the mailbox. |
| `GRAPH_CLIENT_ID` | Microsoft CLI app | Your own app registration, if the tenant needs one. |
| `GRAPH_TENANT_ID` | `organizations` | Your tenant id is a little faster. |
| `GRAPH_SCOPES` | `Mail.ReadWrite,MailboxSettings.Read` | Delegated scopes. |
| `SUBJECT_REGEX` / `HEADER_REGEX` | see above | Request-number patterns, steps 1/3 and 4. |
| `MENTION_REGEX` | `\byou were mentioned\b` | Subject of a mention mail. Empty: mentions are treated like any other mail. |
| `GRAPH_TOKEN_CACHE` | `.state/msal_cache.json` | Refresh token. **Treat it like a password.** |
| `STATE_FILE` | `.state/state.json` | Request-number cache, audit trail, last run. |
| `LOG_DIR` | `logs` | Technical log, `sync.log`, rotating. |
| `DAILY_LOG_DIR` | `logs/daily` | One human-readable file per day. |

## Command line

| Flag | Effect |
|---|---|
| `--apply` | Actually write the categories (default: dry run). |
| `--days N` | Only mails of the last N days (default: no limit). |
| `--limit N` | With `--apply`: change at most N mails. |
| `--check` | Validate configuration and access, then exit. |
| `--explain "text"` | Say why mails matching `text` in the subject were or were not touched. Read-only. |
| `--no-login` | Never open a sign-in window; exit with code 2 instead. For scheduled runs. |
| `--no-cache` | Ignore the request-number cache and parse every candidate again. |
| `--verbose` | Debug logging on stdout. |

Exit codes: `0` success · `1` partial failure or unexpected error · `2` configuration or
authentication problem.

## Running it on a schedule (Windows)

Sign in once in a terminal (`python sync.py --check`), then:

```bash
schtasks /Create /TN "Outlook-Xurrent-Sync" /SC MINUTE /MO 15 /IT /RL LIMITED /F /TR "\"%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe\" \"C:\path\to\outlook-xurrent-sync\sync.py\" --apply --no-login"
```

| Part | Why |
|---|---|
| `/SC MINUTE /MO 15` | every 15 minutes |
| `/IT` | runs with your interactive token, no password is stored, and your token cache is reachable |
| `/RL LIMITED` | no elevated rights needed |
| `pythonw.exe` | no console window flashing up every 15 minutes |
| `--no-login` | if the cached sign-in ever expires, the run ends with exit code 2 and a line in `logs/sync.log` instead of opening a sign-in window nobody sees, which would block the task for good |

Start with the same command *without* `--apply`: a day of `WOULD-APPLY` lines in
`logs/daily/` shows what it would have done. The task only runs while you are signed in.
The task scheduler does not start a second instance while one is running; do not start a
manual run at the same moment either.

`sync-vorschau.cmd` (dry run) and `sync-anwenden.cmd` (apply) are double-click launchers
for the same thing.

## Local files

Created at run time, all in `.gitignore`:

- `.state/msal_cache.json`: Microsoft refresh token, **plain text**. Protect the disk.
- `.state/state.json`: request-number cache and an audit trail of every change with the
  categories before and after (newest 5000 entries).
- `logs/sync.log`: technical log, rotates (1 MB × 5). An unexpected crash is logged here
  with its traceback.
- `logs/daily/YYYY-MM-DD.log`: what was recategorised that day (`APPLIED`, `WOULD-APPLY`,
  `ERROR`, one `RUN` line per run), tab separated and greppable.

These files contain mail subjects. Do not publish them, and keep the project folder out of
OneDrive and similar: a synced `.git` can be damaged by concurrent access, and `.env`
would end up in the cloud.

## Development

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

The tests need no network, no Microsoft account and no Xurrent account.

```
sync.py                  entry point
outlook_xurrent/
  config.py              constants, .env loading, validation
  extract.py             request-number extraction (pure)
  decide.py              done/waiting decision and category rewriting (pure)
  graph.py               Microsoft Graph: sign-in, folders, mails, PATCH categories
  xurrent.py             Xurrent REST, GET only, batch lookup via ?id=1,2,3
  httputil.py            retries, 429/Retry-After, error types
  state.py               JSON state: cache, audit trail, last run
  daily.py               daily journal
  report.py              table, counters, list of mails without a number
  cli.py                 run order and exit codes
tests/                   pytest, fixtures built from an anonymised notification mail
```

## Known limits

- Only the Inbox and its subfolders are scanned.
- The server-side filter on the open category (`categories/any(c:c eq '…')`) is not
  promised by the Graph documentation. It works in the tenant this was built against; if
  Graph rejects it, the script lists the whole folder and filters locally.
- Categories are written as a whole list taken from the moment the mail was listed. If you
  change the categories of that same mail in the few seconds a run takes, the run
  overwrites your change.
- The default `MENTION_REGEX` matches the English Xurrent notification. If your
  notifications are in another language, set `MENTION_REGEX`; the rest of the detection
  does not depend on the language.

## License

[MIT](LICENSE)
