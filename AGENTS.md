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
| Owners, the board, funding, revenue and costs by year, the kinds of slot, the way home, the legs flown per year, the sales books, the competitors, the stage probabilities of the sales forecast | `finance.yaml` |
| Today's date, the history in one line a year, the owner of each drive folder, the headcount of the hangar and the stations | `company.yaml` |
| Ships, classes, crews and their members, the cabin plan, and `pattern`: every leg a ship flies in a month besides the scheduled ones | `fleet.yaml` |
| Ports, handlers, contract dates | `ports.yaml` |
| Seat classes and what travellers are told | `cabin.yaml` |
| Staff, channels | `people.yaml`, `channels.yaml` |
| Corporate accounts, and every deal since January 2098 except the third quarter of 2099 (that one is the win-loss pack's) | `crm/accounts.yaml`, `crm/deals.yaml` |
| The term of every won deal, and the charter orders in the drive | `crm/contracts.yaml` |
| The account plans and the customer stories in the drive | `crm/account-plans.yaml`, `crm/case-studies.yaml` |
| Charter revenue by quarter | `finance.yaml` (`charter_by_quarter`) |
| Charter prices | The rate card, written by `scripts/routine.py` from the snapshot's base prices |
| Bookings, quotes, flights, travellers | `snapshot/*.csv`, written by `scripts/build.py` |

`content/company/drive/ops/how-we-fly.md` explains in the company's words how a small booking snapshot fits a larger company: the four kinds of flying, the three kinds of slot, how a traveller gets home, a month of every ship, who owns what, and why the booking system starts on 10 August. Keep it in step with `finance.yaml` and with `pattern` in `fleet.yaml`.

### How the company hangs together

These hold everything else up. Check a new document against them.

- **A slot decides who may be on board.** A public slot sells seats on the site, and the company holds one a month on each of the ten seed routes. A contract slot carries travellers under a corporate contract. A technical slot carries nobody who paid for that leg. That is why ships fly legs that are not for sale, and why more public departures are a question for the Slot Conference (next: 23 November 2099).
- **The seed has no way home from Luna, Jupiter and Pluto.** The fare includes it: travellers ride back on the Luna shuttle or on the positioning leg (Jupiter on the 19th, Pluto on the 24th), booked through Support on the passage list, which is not in the booking system. From Mars, Venus and Europa the way home is a scheduled departure.
- **A ship is where its next leg starts.** `pattern` in `fleet.yaml` gives every leg of the month. A mail or a log entry about a flight on a given day has to be a leg of the pattern or a scheduled departure of the snapshot. The Luna shuttle docks at Tycho at 01:50 and leaves at 03:20 on ten mornings a month; Tycho counts fourteen turnarounds a month.
- **Crews have rest and a limit of three outer-system legs in thirty days.** Outer A is at the limit.
- **A deal is contract flying and never a public seat.** Its amount is seats times the rate card, less the discount, plus port fees at cost. Deal sizes are small in seats: a seat to Mars is $1,000,000. Where a note names seats, the arithmetic has to work.
- **A won deal is a contract with a term.** `crm/contracts.yaml` gives the first and last month of travel of every won deal, the pack's too. The register, the account plans and the charter revenue by quarter are written from it. The build fails if a won deal has no term, if a charter order does not add up to its deal, or if the contracts of a quarter are more than the charter revenue of `finance.yaml`.
- **The deal history has a shape, and new deals should keep it.** In 2098 most lost deals were lost on price. Since June 2099, when Perihelion went into refit and the Belt call went from two a month to one, most are lost on frequency. No contract departure flies to Pluto: a customer books the scheduled departure on the site and buys reserved seats home. An account's first won deal is in its `since` year, and a deal's owner worked here on its close date (Bastien Moreau from April 2099).
- **Money has one scale.** A laptop is about $79,000, a day's allowance on Luna $9,000, an economy seat to the Moon $500,000, a leg to Mars about $2.7 million to fly (`drive/finance/analysis/cost-per-leg-2099.md`).
- **The snapshot has no booking or quote after 31 October.** The story's reason is the changeover of the ports' traveller registry: no carrier could book from 1 November to 5 November, 06:00 UTC.
- **`people.yaml` holds three kinds of people.** Office staff have no `kind` and are the only ones in the generated office routine. `kind: crew` are the 21 flight crew, `kind: station` the heads of the hangar and the stations. Their teams are counted in `company.yaml`, not named.

### Content

- Content is plain files under `content/`. Prefer adding a file over adding code. Hand-written content goes in `content/company/`, scenario content in `content/packs/`.
- Hand-written content must agree with the snapshot and with the three ground truths. Before stating a number about bookings, flights or quotes, look it up in `snapshot/`.
- Outside the packs, do not add deals that close between 1 July and 30 September 2099, or tickets and mail about price complaints.
- The notes of the win-loss pack's deals are evidence for its ground truth. Do not add to them or reword the reasons.
- Nothing outside the pack counts or explains the wins and losses of the third quarter of 2099: no tracker, no summary by quarter that includes it, no account plan that gives the reason. A won deal of that quarter may appear as a signed contract.
- A customer story names no price, no discount, no competitor and no lost deal.
- The mails and tickets of the quote-mismatch pack name bookings 309, 335 and 338. Check them after any change to the snapshot rules.

### Generated files

Run `.venv/bin/python scripts/build.py`. It is deterministic and rewrites all of these; never edit them by hand.

- `snapshot/*.csv` and the packs' `ground-truth.md` and exports: `scripts/build.py`.
- `content/generated/` (mail, chat, calendar, tickets, drive): wiped and rewritten by `scripts/routine.py`.
- Office files (.xlsx, .pptx, .docx): written in the house style of `scripts/office.py` (`Book`, `Deck`, `Doc`). `scripts/documents.py` holds the general ones, `scripts/worldbook.py` those about ships, ports and the cabin, `scripts/salesbook.py` the contract register, the charter orders, the customer stories, the account plans, the quarterly trackers and the charter revenue by quarter, `scripts/companybook.py` the company profile, the financial summary and `docs/company.html`.
- `content/previews/`: a PDF of every deck and Word document, rendered with LibreOffice for changed files only. Commit them with the office files. Without LibreOffice the build removes the pages of changed files; run `scripts/previews.py` on a machine that has it.

A new generated document takes booking and deal figures from the snapshot and the CRM, repeats figures from existing notes exactly, is dated on or before the story's today, does not answer a question that a pack or an open mail asks, and has an `owner` and a `modified` date. The docstring of `documents.py` lists what the documents may not say.

### The server and the web view

- `backoffice/store.py` loads content and state from files on every call, with no cache.
- `backoffice/server.py` defines the MCP tools. Register a new one with `@_read(system)` or `@_write(system)`; a write tool must go through `@_write` so that it is permission-checked and logged.
- `crm_pipeline_summary` totals open deals only. Do not extend it, or add another tool, to total won and lost deals by period: that would count the third quarter of 2099, which only the win-loss pack may do.
- Tools run in worker threads and write tools share one lock. The text of drive files is cached by path, size and modification time; nothing else is.
- `backoffice/web.py` and `backoffice/static/index.html` are the web view. It follows the layouts people know (three-pane mail, a channel sidebar, a folder tree) but uses no names, logos or artwork of real products. Systems are named generically: mail, chat, calendar, CRM, tickets, drive.
- To verify a change: `./start.sh`, open http://localhost:8090, and call the tools from an MCP client. There is no test suite yet.
