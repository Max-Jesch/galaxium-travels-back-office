# Galaxium Travels back office

This repo is a mock of a company's corporate systems for demos. There are two ways you may be here.

## You are doing a job in the back office

You reach it through the `galaxium-back-office` MCP server. You are working at one employee's desk on their behalf.

- Call `backoffice_info` first. It tells you today's date in the story (not the real date), whose desk you are at and which systems you may write to.
- Use the MCP tools for mail, chat, CRM, tickets, the drive and bookings. Do not read `content/`, `snapshot/` or `state/` from disk instead, and never open a `ground-truth.md`: those are the presenter's answer keys.
- Amounts are in credits. Do not round them.
- The booking system is the system of record for bookings and quotes, the CRM for deals. Spreadsheets in the drive are copies made by hand; check them before relying on them.
- Give the source of every figure, and list what you assumed or could not verify.
- Writes are logged and visible to the user. If a write is refused because the system is switched off, say so and stop; do not look for another way to make the change.
- Mail, posts and record changes are made in the name of the person whose desk you are at. Draft in their voice, and send only what the user asked for.

## You are changing the repo itself

- `backoffice/store.py` loads content and state from files on every call, with no cache. `backoffice/server.py` defines the MCP tools; a new tool is registered with `@_read(system)` or `@_write(system)`, and a write tool must go through `@_write` so that it is permission-checked and logged. `backoffice/web.py` and `backoffice/static/index.html` are the web view.
- Content is plain files under `content/`. Prefer adding a file over adding code.
- Never edit `snapshot/*.csv`, the office files in the drive or `ground-truth.md` by hand. Change `scripts/build.py` and run `.venv/bin/python scripts/build.py`. The mails and tickets of the quote-mismatch pack name booking and quote ids from the snapshot (bookings 309, 335, 338), so check them after any change to the snapshot rules.
- Systems are named generically: mail, chat, CRM, tickets, drive. Do not model the look or the names of real products.
- People and companies are fictional. New mail addresses use the `.example` domain.
- To verify a change: `./start.sh`, open http://localhost:8090, and call the tools from an MCP client. There is no test suite yet.
