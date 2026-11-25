# Galaxium Travels back office

This repo is a mock of a company's corporate systems for demos. There are two ways you may be here.

## You are doing a job in the back office

You reach it through the `galaxium-back-office` MCP server. You are working at one employee's desk on their behalf.

- Call `backoffice_info` first. It tells you today's date in the story (not the real date), whose desk you are at and which systems you may write to.
- Use the MCP tools for mail, chat, the calendar, CRM, tickets, the drive and bookings. Do not read `content/`, `snapshot/` or `state/` from disk instead, and never open a `ground-truth.md`: those are the presenter's answer keys.
- To build on a drive file such as the deck template, copy it into your working folder with `drive_download`, work on the copy and put the result on the drive with `drive_upload`. Do not look for drive files on disk.
- Amounts are in dollars. Do not round them.
- The booking system is the system of record for bookings and quotes, the CRM for deals. Spreadsheets in the drive are copies made by hand; check them before relying on them.
- Give the source of every figure, and list what you assumed or could not verify.
- Writes are logged and visible to the user. If a write is refused because the system is switched off, say so and stop; do not look for another way to make the change.
- Every write tool returns a `url` that opens the record in the web view. When you report what you did, link each thing you created or changed with it. `backoffice_info` gives the link pattern for records you only read.
- Mail, posts and record changes are made in the name of the person whose desk you are at. Draft in their voice, and send only what the user asked for.

## You are changing the repo itself

Read `docs/company.html` first. It is the company on one page: how it earns, what it costs, who owns it, how it flies. The build writes it, so never edit it.

### The setting

- The story's today is Wednesday 4 November 2099 (`content/company/company.yaml`). Nothing is dated after it, except plans.
- Amounts are US dollars. "Service credit" and "travel credit" are contract and refund terms, not a currency.
- People and companies are fictional. New mail addresses use the `.example` domain.
- Places have two names and both are right: the booking app names the body and the destination (Moon, Olympus Base), the company names the port (Luna, Tycho Spaceport, Olympus Port). `ports.yaml` holds the pairs; do not rename either side.

### Where a fact lives

Change a fact in its file, then run the build. Do not type it into a script.

| Fact | File in `content/company/` |
| --- | --- |
| Owners, funding, revenue and costs by year, the flying programme, the sales books, the competitors | `finance.yaml` |
| Today's date, the history in one line a year, the owner of each drive folder | `company.yaml` |
| Ships, classes, crews, the cabin plan | `fleet.yaml` |
| Ports, handlers, contract dates | `ports.yaml` |
| Seat classes and what travellers are told | `cabin.yaml` |
| Staff, channels, corporate accounts, deals | `people.yaml`, `channels.yaml`, `crm/` |
| Bookings, quotes, flights, travellers | `snapshot/*.csv`, written by `scripts/build.py` |

`content/company/drive/ops/how-we-fly.md` explains in the company's words how a small booking snapshot fits a larger company: the three kinds of flying, who owns what, slots, and why the booking system starts on 10 August. Keep it in step with `finance.yaml`.

### Content

- Content is plain files under `content/`. Prefer adding a file over adding code. Hand-written content goes in `content/company/`, scenario content in `content/packs/`.
- Hand-written content must agree with the snapshot and with the three ground truths. Before stating a number about bookings, flights or quotes, look it up in `snapshot/`.
- Outside the packs, do not add deals that close between 1 July and 30 September 2099, or tickets and mail about price complaints.
- The mails and tickets of the quote-mismatch pack name bookings 309, 335 and 338. Check them after any change to the snapshot rules.

### Generated files

Run `.venv/bin/python scripts/build.py`. It is deterministic and rewrites all of these; never edit them by hand.

- `snapshot/*.csv` and the packs' `ground-truth.md` and exports: `scripts/build.py`.
- `content/generated/` (mail, chat, calendar, tickets, drive): wiped and rewritten by `scripts/routine.py`.
- Office files (.xlsx, .pptx, .docx): written in the house style of `scripts/office.py` (`Book`, `Deck`, `Doc`). `scripts/documents.py` holds the general ones, `scripts/worldbook.py` those about ships, ports and the cabin, `scripts/companybook.py` the company profile, the financial summary and `docs/company.html`.
- `content/previews/`: a PDF of every deck and Word document, rendered with LibreOffice for changed files only. Commit them with the office files. Without LibreOffice the build removes the pages of changed files; run `scripts/previews.py` on a machine that has it.

A new generated document takes booking and deal figures from the snapshot and the CRM, repeats figures from existing notes exactly, is dated on or before the story's today, does not answer a question that a pack or an open mail asks, and has an `owner` and a `modified` date. The docstring of `documents.py` lists what the documents may not say.

### The server and the web view

- `backoffice/store.py` loads content and state from files on every call, with no cache.
- `backoffice/server.py` defines the MCP tools. Register a new one with `@_read(system)` or `@_write(system)`; a write tool must go through `@_write` so that it is permission-checked and logged.
- `backoffice/web.py` and `backoffice/static/index.html` are the web view. It follows the layouts people know (three-pane mail, a channel sidebar, a folder tree) but uses no names, logos or artwork of real products. Systems are named generically: mail, chat, calendar, CRM, tickets, drive.
- To verify a change: `./start.sh`, open http://localhost:8090, and call the tools from an MCP client. There is no test suite yet.
