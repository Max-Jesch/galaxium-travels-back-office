# Galaxium Travels back office

The corporate systems of [Galaxium Travels](https://github.com/IBM/galaxium-travels), the fictional interplanetary booking company: mail, chat, CRM, tickets, a shared drive and a snapshot of bookings and quotes. Everything is a mock, served over MCP from one process on your own computer, so that a desktop agent such as Bob Desktop has an office to work in.

No accounts, no API keys, and nothing leaves your machine.

## Start it

You need Python 3.10 or newer.

```bash
./start.sh
```

The first start sets itself up, which takes about a minute and needs the internet once. After that:

- **Web view:** http://localhost:8090 shows the company the way an employee sees it: a three-pane inbox, chat with channels, threads and direct messages, a calendar, a drive with folders and file previews, a ticket queue, the CRM pipeline, the booking system and the staff directory. Everything the agent writes is marked in green where it lands, and listed on the *Activity* tab.
- **MCP server:** http://localhost:8090/mcp

To return to the starting state after a demo, press **Reset** in the web view or run:

```bash
./reset.sh
```

## Connect an agent

Point any MCP client at `http://localhost:8090/mcp` (streamable HTTP). `.bob/mcp.json` registers it for Bob in this folder.

If your client starts MCP servers itself instead of connecting to a URL, use the stdio mode. The web view still shows what happens, as long as `./start.sh` is running too:

```json
{
  "mcpServers": {
    "galaxium-back-office": {
      "command": "/full/path/to/galaxium-travels-back-office/.venv/bin/python",
      "args": ["-m", "backoffice", "--stdio"],
      "cwd": "/full/path/to/galaxium-travels-back-office"
    }
  }
}
```

## Scenarios

A scenario pack is one corporate job with the mail, chat, records and files it needs. Choose it in the web view, or start with `./start.sh --pack win-loss`. Only the company's everyday content and the chosen pack are loaded, so the stories do not get in each other's way.

| Pack | Department | The job | Whose desk |
| --- | --- | --- | --- |
| `quote-mismatch` | Finance | Customers were charged a different price than they were quoted. Reconcile October's quotes against bookings, write the memo, hand the cause to engineering. | Tomás Rivera, Financial Analyst |
| `win-loss` | Sales | Why did we win and lose the third quarter's corporate deals? Analyse the CRM, check the tracker spreadsheet, build the deck. | Amara Okafor, Head of Sales |
| `vendor-contract` | Legal | Review the spaceport vendor's renewal draft against the contract playbook and prepare the escalation. | Sam Whitlock, Contracts Manager |

Each pack has:

- a **prompt** to give the agent, shown on the *Presenter check* tab,
- a **ground truth** with the expected numbers and findings, on the same tab and in `content/packs/<pack>/ground-truth.md`,
- **one planted error** in the source material that a careful result catches,
- **one claim by a colleague** that the data does not support.

The agent cannot read the ground truth through MCP. Keep the *Presenter check* tab off screen until you check the result.

The company content alone (choose *Company only*) supports two more jobs without a ground truth yet: preparing Monday's route review from mail, chat and the drive, and the weekly bookings report from `templates/weekly-bookings-report-template.docx`.

## What the agent can do

29 tools, named by system. Reading is open. The eight writing tools are marked as such, can be switched off per system on the *Permissions* tab or in `permissions.yaml`, and every call to one is logged on the *Activity* tab with its arguments and outcome, including calls that were refused.

| System | Read | Write |
| --- | --- | --- |
| Mail | `mail_list_threads`, `mail_read_thread` | `mail_send` |
| Chat | `chat_list_channels`, `chat_read_channel`, `chat_search` | `chat_post_message` |
| Calendar | `calendar_list_events` | `calendar_create_event` |
| Tickets | `tickets_list`, `tickets_get` | `tickets_create`, `tickets_update` |
| CRM | `crm_list_accounts`, `crm_list_deals`, `crm_get_deal` | `crm_update_deal`, `crm_add_account_note` |
| Drive | `drive_list`, `drive_read`, `drive_search` | `drive_write_text` |
| Bookings | `bookings_list_flights`, `bookings_find_customers`, `bookings_search`, `bookings_summary`, `bookings_search_quotes` | none |
| Directory | `backoffice_info`, `directory_list_people` | none |

Mail is never really sent. Writes go to the `state/` folder and nowhere else.

The drive holds real `.xlsx`, `.docx` and `.pptx` files. `drive_read` returns them as text, and `local_path` in the result is the file itself for agents that can open office files.

## Add your own content

Everything the back office shows is a plain file under `content/`. Add a file and it appears on the next click; no restart and no code.

```
content/
  company/               always loaded
    people.yaml          the cast: 48 employees in ten departments
    customers.yaml       the ten travellers from the booking app's seed data
    channels.yaml        channels by sidebar section, and the direct messages
    mail/*.md            one file per thread
    chat/<channel>/*.md  one file per day, or one file with several days
    calendar/*.yaml      meetings and deadlines
    tickets.yaml
    crm/accounts.yaml, crm/deals.yaml
    drive/               the shared drive, as real folders and files
  generated/             same layout: the daily routine, written by scripts/build.py
  packs/<pack>/          same layout, plus pack.yaml and ground-truth.md
snapshot/                flights, users, bookings and quotes as CSV
```

A mail thread looks like this. Messages are separated by a line with three dashes:

```markdown
---
subject: Cancelling my Europa trip, what do I get back?
labels: [support, refund]
---
From: Heidi <heidi@europa.com>
To: Galaxium Support <support@galaxium.example>
Date: 2026-11-03 19:02

Hello, I may have to cancel my trip.
```

A mail thread may also name `attachments: [path/in/the/drive.xlsx]` and `flagged: true` in its header.

A chat file holds one or more days. An indented message is a reply in the thread of the message above it, and an indented line starting with `+` gives reactions:

```markdown
## 2026-11-02
[09:30] Bruno Ferreira: Multi-factor sign-in becomes mandatory on Monday.
    [09:44] Hana Kobayashi: Does that include the payroll system?
    + 👍 4
```

A direct message is a channel with `type: dm` and `members` in `channels.yaml`. Only its members see it, in the web view and through the tools.

A drive file shows an owner and a date in the web view. A markdown file can state them in a header (`owner:`, `modified:`); other files get the owner of their top folder from `company.yaml`.

### How much is in there

About 180 mail threads, 500 chat messages in 32 channels, 200 tickets, 100 drive files, 265 calendar entries, 28 corporate accounts and 30 to 48 deals, depending on the pack. Roughly a third is written by hand under `content/company/`: the conversations, decisions and documents. The rest is routine that a real company produces without thinking: daily booking summaries, vendor invoices, customer questions, deployment notices, weekly reports and recurring meetings. That part is generated from the snapshot by `scripts/routine.py`, so its numbers agree with the booking system.

`content/company/drive/ops/how-we-fly.md` explains how the small booking snapshot (one scheduled departure per route per month) fits a company with deals worth tens of millions: charter flying is sold through the CRM and is not in the booking system.

In the story, today is always **Wednesday 4 November 2026** (`content/company/company.yaml`). Write dates relative to that.

## How it fits the booking app

The back office does not need the booking app to run. It carries its own fixed snapshot, so that the numbers in mails, tickets and reports always agree:

- 60 flights: the ten routes and base prices of the booking app's seed data, with one departure a month from September 2026 to February 2027. Flight ids 1 to 10 are the seed's routes in the seed's order.
- 60 travellers: the seed's ten (Alice to Judy, same ids) and fifty more.
- 340 bookings from August to October 2026, and 238 price quotes from October, when the hold service went live.

The quote mismatch is real behaviour of the booking app, not an invention: its Java hold service prices a quote from its own table and ignores the route, and the Python backend charges base price times seat class factor at booking. The snapshot applies both rules exactly as the code does.

The snapshot, the office files and the ground truths are generated by `scripts/build.py` and committed. Run it again only if you change its rules or the win-loss deals; it always produces the same numbers.

## Also in this repo

- `AGENTS.md`: guidance for an agent working in the back office.
- `skills/`: two example skills that package a team's method, one for reconciling two sources and one for the weekly report, which is written to run on a schedule.
- `Galaxium Back Office Repo Design.md`: why this repo exists and the decisions behind it.

All people and companies are fictional. Employee and vendor addresses use the reserved `.example` domain; the ten travellers keep the addresses they have in the booking app's seed data.
