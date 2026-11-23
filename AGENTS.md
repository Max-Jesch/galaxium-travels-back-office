# Galaxium Travels back office

This repo is a mock of a company's corporate systems for demos. There are two ways you may be here.

## You are doing a job in the back office

You reach it through the `galaxium-back-office` MCP server. You are working at one employee's desk on their behalf.

- Call `backoffice_info` first. It tells you today's date in the story (not the real date), whose desk you are at and which systems you may write to.
- Use the MCP tools for mail, chat, the calendar, CRM, tickets, the drive and bookings. Do not read `content/`, `snapshot/` or `state/` from disk instead, and never open a `ground-truth.md`: those are the presenter's answer keys.
- Amounts are in credits. Do not round them.
- The booking system is the system of record for bookings and quotes, the CRM for deals. Spreadsheets in the drive are copies made by hand; check them before relying on them.
- Give the source of every figure, and list what you assumed or could not verify.
- Writes are logged and visible to the user. If a write is refused because the system is switched off, say so and stop; do not look for another way to make the change.
- Every write tool returns a `url` that opens the record in the web view. When you report what you did, link each thing you created or changed with it. `backoffice_info` gives the link pattern for records you only read.
- Mail, posts and record changes are made in the name of the person whose desk you are at. Draft in their voice, and send only what the user asked for.

## You are changing the repo itself

- `backoffice/store.py` loads content and state from files on every call, with no cache. `backoffice/server.py` defines the MCP tools; a new tool is registered with `@_read(system)` or `@_write(system)`, and a write tool must go through `@_write` so that it is permission-checked and logged. `backoffice/web.py` and `backoffice/static/index.html` are the web view.
- Content is plain files under `content/`. Prefer adding a file over adding code. Hand-written content goes in `content/company/`. `content/generated/` is wiped and rewritten by `scripts/routine.py` (run through `scripts/build.py`): never edit it by hand.
- Hand-written content must agree with the snapshot and with the three ground truths. Before stating a number about bookings, flights or quotes, look it up in `snapshot/`. Do not add deals that close between 1 July and 30 September 2026, or tickets and mail about price complaints, outside the packs. `content/company/drive/ops/how-we-fly.md` is the fiction that reconciles the small snapshot with a larger company. Places have two names and both are right: the booking app names the body and the destination (Moon, Olympus Base), the company names the port (Luna, Tycho Spaceport, Olympus Port). `content/company/ports.yaml` holds the pairs; do not rename either side.
- Office files (.xlsx, .pptx, .docx) are written by code, in the house style of `scripts/office.py` (`Book`, `Deck`, `Doc`). The decks, workbooks and documents of the generated drive are in `scripts/documents.py`; its docstring lists what they may and may not say. The documents about the ships, the ports and the cabin are in `scripts/worldbook.py` and are written from `content/company/fleet.yaml`, `ports.yaml` and `cabin.yaml`: change a fact about a ship, a port or a seat class in those files, not in the code. The fleet register and the ships and crews of the departure plan read `fleet.yaml` too. A new one must take booking and deal figures from the snapshot and the CRM file, repeat figures from existing notes exactly, be dated on or before 4 November 2026, and not answer a question that a pack or an open mail asks. Give it an `owner` and a `modified` date; they end up in `content/generated/drive/_index.yaml`.
- `content/previews/` holds a PDF of every deck and Word document, for the preview in the web view. `scripts/build.py` renders them with LibreOffice, and only for files that changed. Commit them together with the office files. Without LibreOffice the build still runs but removes the pages of changed files; run `scripts/previews.py` on a machine that has it.
- Never edit `snapshot/*.csv`, the office files in the drive or `ground-truth.md` by hand. Change `scripts/build.py` and run `.venv/bin/python scripts/build.py`. The mails and tickets of the quote-mismatch pack name booking and quote ids from the snapshot (bookings 309, 335, 338), so check them after any change to the snapshot rules.
- Systems are named generically: mail, chat, calendar, CRM, tickets, drive. The web view follows the layout conventions people know from popular tools (three-pane mail, a channel sidebar, a folder tree), so that it reads as familiar at a glance. It does not use the names, logos or artwork of real products.
- People and companies are fictional. New mail addresses use the `.example` domain.
- To verify a change: `./start.sh`, open http://localhost:8090, and call the tools from an MCP client. There is no test suite yet.
