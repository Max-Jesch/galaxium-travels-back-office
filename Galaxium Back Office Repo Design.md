# Galaxium Back Office: Repo Design

Nov 16, 2026 · @Max

## Why a separate repo

The back office is a mock of Galaxium's corporate systems, served over MCP, that anyone can start with one command. It lives in its own repo so that someone demoing Bob Desktop does not need the Java and Node toolchain of `galaxium-travels` just to get an inbox.

It lets the Bob Desktop stories show what the booking repo cannot: gathering from several sources, tracking records and posting somewhere.

## Context

IBM is about to launch Bob Desktop, a desktop agent for knowledge workers in the same category as the Claude app and ChatGPT Work. The back office is the demo environment for the YouTube and tutorial content about it.

- **The existing demo.** [IBM/galaxium-travels](https://github.com/IBM/galaxium-travels) is a fictional interplanetary booking company, built as three services to show coding agents at work. It is self-contained in one repo, which is why it is easy for others to replicate.
- **The new story.** The series goes behind the scenes at Galaxium: one corporate job per department, done with Bob Desktop, with the check of the result shown on camera.
- **Who runs this repo.** Presenters, sellers and viewers following along. Most are not developers, so setup has to be one command and the content has to be readable as plain files.
- **What it must make visible.** Gathering from several sources, tracking records and posting somewhere. Those are the things a single code repo cannot show.

## The jobs the back office must support

Six recurring corporate jobs cover what competitors and usage data describe. Each scenario pack should serve at least one of them.

| Job | Galaxium example | What the back office must provide |
| --- | --- | --- |
| Gather and brief | Ops lead preparing the Monday route review | Related threads spread across mail, chat and the drive |
| Recurring report | Weekly bookings and cancellations report | A report template and fresh data behind it each week |
| Numbers work | Finance month-end: quotes against bookings | Exports that can be reconciled against the snapshot |
| Decide and present | Sales win-loss, or the case for a new Pluto route | CRM deals with outcomes, call notes, a deck template |
| Review against a standard | Legal reviewing a spaceport vendor contract | A contract, a playbook and vendor proposals in the drive |
| Keep the records current | Support filing and routing complaints | Writable tickets and CRM records, a channel to post to |

Two needs cut across all six: a job should be repeatable on a schedule, and a team's method should be packageable as a skill.

## Research findings that steer the build

Findings are from research done on 16 Nov 2026. Each one maps to a build decision.

| Finding | What it means for the repo |
| --- | --- |
| Most real usage is "the work around the work". In Anthropic's sample of 1.2 million Cowork sessions (May 2026), business operations was 33.4% and content creation 16.4%; data analysis was 5.8%. Vendor data, classified automatically. ([source](https://claude.com/blog/how-people-are-using-claude-cowork)) | Fill the back office with status reports, trackers, threads and decks first. Heavy analytics is a minority case. |
| Every competitor promises the finished deck or sheet. None teaches how to check it. Anthropic's guide comes closest, with tips such as asking for citations and for a list of guesses. ([source](https://claude.com/resources/guides/claude-cowork-product-guide/popular-use-cases)) | Every scenario needs a ground truth that can be checked on camera: numbers that reconcile to the snapshot, sources that exist as files, and one planted error. |
| Product surface names are short-lived. Anthropic merged Cowork into chat and OpenAI merged the Codex app into ChatGPT desktop within 2026. | Name systems generically: mail, chat, CRM, tickets, drive. No lookalikes of real products. |
| Governance of connectors is the durable enterprise topic, and it is Bob's positioning. | Keep read and write tools separate, log every write, and make permissions visible in the web view. |
| MCP, Agent Skills and AGENTS.md now sit under the Linux Foundation's Agentic AI Foundation (Agent Skills joined 22 Oct 2026). | Build on plain MCP. Ship an `AGENTS.md` and example skills in the repo. |
| IBM's own TechXchange demo was a sales win-loss analysis from a tracker board and an Excel file, ending in a presentation. OpenAI's flagship example is finance month-end close. ([IBM coverage](https://codezine.jp/article/detail/30219), [OpenAI](https://openai.com/index/chatgpt-for-your-most-ambitious-work/)) | The first packs should cover sales and finance. The drive needs real `.xlsx` and `.pptx` files, not only markdown. |
| IBM reports about 170,000 internal Bob users, of whom about 120,000 are non-developer knowledge workers. ([source](https://enterprisezine.jp/article/detail/25402)) | The audience is mostly non-technical. Setup and content must not assume a developer. |

No independent study of non-developers using agent desktops was found. Demo content should not cite productivity figures.

## Facts from galaxium-travels to reuse

Read from `main` at commit `9c5cbaa`. The pricing finding comes from reading the code; the app was not run.

- **Stack to match.** The booking backend is Python with FastAPI and FastMCP, and serves MCP at `:8001/mcp`. `.bob/mcp.json` is where the project's MCP servers are registered.
- **Data model.** `users`, `flights` and `bookings`. A booking has a seat class, a status (`booked`, `cancelled`, `completed`) and a `price_paid`.
- **Customers.** The seed has ten travellers, Alice to Judy, with addresses such as `alice@example.com` and `eve@marsmail.com`. Reuse them as the customers in mail threads and tickets.
- **Routes and prices.** Ten routes with a base price each, from Earth to Moon at 500,000 to Earth to Pluto at 5,000,000. Seat classes multiply the base price: economy 1, business 2.5, galaxium 5.
- **The seed is not reproducible.** `seed.py` uses unseeded `random` and creates only 20 bookings. The back office snapshot must be deterministic and hold a few hundred bookings and quotes.
- **Business rules worth a story.** Holds expire after 15 minutes. A booking is rejected when the name does not match the user id. Issue #43 adds a refund preview on cancel.

### The quote mismatch

The Java hold service prices quotes with its own table, and the Python backend recalculates the price at booking and ignores the quote. The Java table knows `economy`, `business` and `first`, so `galaxium` falls back to the economy price.

| Earth to Mars | Quoted by Java `PricingService` | Charged by Python `book_flight` |
| --- | --- | --- |
| Economy | 550,000 | 1,000,000 |
| Business | 2,750,000 | 2,500,000 |
| Galaxium | 550,000 | 5,000,000 |

This is the strongest candidate for the first scenario: a customer complaint, a reconciliation of quotes against bookings, a memo, and a ticket handed to engineering.

## The repo

| Part | What it is |
| --- | --- |
| One process | A Python MCP server, matching the FastMCP pattern in the booking backend, plus a small web page that shows the inbox, channels and ticket queue. One command to start, no accounts. |
| Four systems | Mail, chat, CRM and tickets, plus a shared drive folder for the spreadsheets, policies and deck templates. |
| Content as plain files | Each mail thread, chat conversation and ticket is a markdown or YAML file under `content/`. Anyone can add a conversation with a pull request, and the company's story is readable without running anything. |
| Scenario packs | Each episode gets a folder with the threads, records and files it needs, so the win-loss story and the month-end story do not pollute each other. |
| Read and write tools kept apart | Reading is open. Sending a mail, posting to a channel or updating a record writes to a local state folder. A reset command restores the starting state. |
| A fixed cast | A dozen named employees across finance, sales, support, legal and ops, with the travellers from the booking seed data as the customers. |

## Design decisions

**Booking data ships as a fixed snapshot.** The back office carries its own snapshot of bookings and quotes, so the numbers in mails and reports always agree. Connecting to the live booking API stays optional.

**A real tenant comes later, without a redesign.** Because the content is plain files, a later exporter can push the same threads and records into Google Workspace or Microsoft 365.

## Build requirements and non-goals

**Requirements**

- One command starts everything. No accounts, no API keys, no network access needed at run time.
- All data is deterministic. Two people who start the repo see the same inbox, records and numbers.
- A reset command returns the back office to its starting state after a demo has written to it.
- Content is authored as plain files and loaded at start. Adding a thread or a ticket needs no code change.
- Each scenario pack includes a ground-truth file with the expected answers, so a presenter can check the agent's output.
- Every write through MCP is logged with the tool, the arguments and the time, and shows up in the web view.
- The drive holds real office files (`.xlsx`, `.docx`, `.pptx`) where a job needs them.
- People and companies are fictional, and mail addresses use fictional or reserved domains.

**Non-goals**

- No clone of Gmail, Slack or Salesforce. The web view only has to show that something arrived or changed.
- No real mail sending, no authentication and no multi-user state.
- No live dependency on the booking app. The snapshot is enough.
- No exporter to Google Workspace or Microsoft 365 yet. Keep the file format clean enough that one can be added.

## Open questions

- [ ] Which story is the first scenario pack: the quote mismatch or the sales win-loss?
- [ ] Can Bob Desktop connect to a local MCP server from a project config, or only to connectors installed from Bazaar?

* [ ] Is the public name of the connector marketplace Bob Bazaar or Bob Marketplace? Both appear in TechXchange material.
* [ ] What is the public repo name and owner: under `IBM/` next to `galaxium-travels`?
