"""The Galaxium back office as one MCP server: mail, chat, CRM, tickets, drive and bookings.

Read tools are open. Write tools go through `_write`, which checks the
permission for that system and logs the call, allowed or not.
"""

from __future__ import annotations

import functools
from collections import defaultdict
from typing import Any, Callable
from urllib.parse import quote

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from . import store

mcp = FastMCP(
    "galaxium-back-office",
    instructions=(
        "The corporate systems of Galaxium Travels, a fictional interplanetary travel company: "
        "mail, chat, CRM, tickets, a shared drive and a read-only snapshot of bookings and quotes. "
        "Call backoffice_info first to learn today's date and whose desk you are working at. "
        "Amounts are in credits. Tools that send, post, create or update are logged, and return a url "
        "that opens the record in the web view: give it to the user when you report what you did."
    ),
    stateless_http=True,
    json_response=True,
)

TOOLS: list[dict] = []  # shown on the Permissions tab of the web view


def _register(fn: Callable, system: str, kind: str) -> Callable:
    TOOLS.append({"name": fn.__name__, "system": system, "kind": kind,
                  "description": (fn.__doc__ or "").strip().split("\n")[0]})
    read_only = kind == "read"
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=read_only, destructiveHint=False, openWorldHint=False))(fn)
    return fn


def _read(system: str) -> Callable:
    return lambda fn: _register(fn, system, "read")


def _write(system: str) -> Callable:
    def decorate(fn: Callable) -> Callable:
        @functools.wraps(fn)
        def guarded(*args: Any, **kwargs: Any) -> Any:
            if not store.permissions().get(system, True):
                store.audit(fn.__name__, system, kwargs, "denied", "Writing is switched off for this system")
                raise PermissionError(
                    f"Writing to {system} is switched off in the back office permissions. "
                    "Nothing was changed. Tell the user; do not try another route."
                )
            try:
                result = fn(*args, **kwargs)
            except Exception as exc:
                store.audit(fn.__name__, system, kwargs, "error", str(exc))
                raise
            summary = result.get("id") or result.get("thread_id") or result.get("path") or result.get("channel") or "ok"
            store.audit(fn.__name__, system, kwargs, "ok", str(summary))
            return result
        return _register(guarded, system, "write")
    return decorate


def _linked(result: dict, app: str, *parts: Any) -> dict:
    """Add the address of the record in the web view, for the agent to cite."""
    path = "/".join(quote(str(p), safe="/") for p in (app, *parts))
    return {**result, "url": f"{store.web_url()}/#{path}"}


def _matches(record: Any, query: str) -> bool:
    return all(word in str(record).lower() for word in query.lower().split())


def _public(entry: dict) -> dict:
    """A drive entry without its place on disk: agents fetch the file with drive_download."""
    return {k: v for k, v in entry.items() if k != "local_path"}


def _all(items: list) -> dict:
    """Lists are wrapped in an object, so that every client shows the whole result."""
    return {"total": len(items), "items": items}


def _page(items: list, limit: int, offset: int = 0) -> dict:
    return {"total": len(items), "offset": offset, "items": items[offset:offset + limit]}


# ---------- orientation ----------

@_read("directory")
def backoffice_info() -> dict:
    """Today's date in the back office, whose desk you are at, and what systems exist. Call this first."""
    me = store.persona()
    return {
        "company": store.company().get("name"),
        "today": store.today(),
        "you_are_working_for": {k: me[k] for k in ("name", "role", "department", "email")},
        "currency": "credits",
        "systems": {
            "mail": "Threads in the company mail system",
            "chat": "Team channels and your direct messages",
            "calendar": "Meetings and deadlines",
            "crm": "Corporate accounts and deals",
            "tickets": "Support and engineering queues",
            "drive": "Shared files: policies, templates, exports",
            "bookings": "Read-only snapshot of flights, bookings and price quotes",
        },
        "write_permissions": store.permissions(),
        "web_view": store.web_url(),
        "links": "Write tools return a url for what they changed. Other records open at "
                 "<web_view>/#tickets/<id>, #mail/<thread id>, #chat/<channel>, #drive/<path>, "
                 "#calendar/<event id>, #crm/deal/<id> and #crm/account/<id>.",
    }


@_read("directory")
def directory_list_people(department: str = "") -> dict:
    """List Galaxium employees with role, department, manager, location and mail address."""
    return _all([p for p in store.people() if not department or p["department"].lower() == department.lower()])


# ---------- mail ----------

def _thread_summary(t: dict) -> dict:
    return {"id": t["id"], "subject": t["subject"], "labels": t["labels"], "messages": len(t["messages"]),
            "last_date": t["last_date"], "from": t["messages"][0]["from"] if t["messages"] else "",
            "preview": t["messages"][-1]["body"][:160] if t["messages"] else ""}


@_read("mail")
def mail_list_threads(mailbox: str = "me", query: str = "", label: str = "", limit: int = 25) -> dict:
    """List mail threads, newest first. mailbox is "me" (default), "all", or a person's name or address."""
    threads = store.mail_threads()
    if mailbox != "all":
        person = store.persona() if mailbox == "me" else store.find_person(mailbox)
        addresses = store.mailbox_addresses(person) if person else [mailbox.lower()]
        threads = [t for t in threads if set(addresses) & set(t["participants"])]
    if label:
        threads = [t for t in threads if label in t["labels"]]
    if query:
        threads = [t for t in threads if _matches(t, query)]
    return _page([_thread_summary(t) for t in threads], limit)


@_read("mail")
def mail_read_thread(thread_id: str) -> dict:
    """Read every message in one mail thread."""
    for t in store.mail_threads():
        if t["id"] == thread_id:
            return t
    raise ValueError(f"No mail thread with id {thread_id}")


@_write("mail")
def mail_send(to: str, subject: str, body: str, cc: str = "", thread_id: str = "") -> dict:
    """Send a mail from your desk. Give thread_id to reply in an existing thread. Mail stays inside the back office."""
    sent = store.send_mail(to, subject, body, cc, thread_id)
    return _linked(sent, "mail", sent["thread_id"])


# ---------- chat ----------

@_read("chat")
def chat_list_channels() -> dict:
    """List the chat channels and your direct messages, with purpose and message count."""
    counts: dict[str, int] = defaultdict(int)
    for m in store.chat_messages():
        counts[m["channel"]] += 1
    return _all([{**c, "messages": counts[c["name"]]} for c in store.channels()])


@_read("chat")
def chat_read_channel(channel: str, since_date: str = "", limit: int = 100) -> dict:
    """Read a channel or one of your direct messages, oldest first. Replies sit under the message they answer. since_date is YYYY-MM-DD."""
    channel = channel.lstrip("#")
    names = [c["name"] for c in store.channels()]
    if channel not in names:
        raise ValueError(f"No channel named {channel}. Channels: {', '.join(names)}")
    messages = [m for m in store.chat_messages(channel) if m["date"] >= since_date]
    return _page(messages[-limit:], limit)


@_read("chat")
def chat_search(query: str, limit: int = 30) -> dict:
    """Search all channels for messages containing every word of the query."""
    return _page([m for m in store.chat_messages() if _matches(m, query)], limit)


@_write("chat")
def chat_post_message(channel: str, text: str) -> dict:
    """Post a message to a channel or one of your direct messages, as the person whose desk you are at."""
    posted = store.post_chat(channel, text)
    return _linked(posted, "chat", posted["channel"])


# ---------- calendar ----------

def _attends(event: dict, person: dict) -> bool:
    names = [event.get("organizer", "")] + list(event.get("attendees", []))
    return any(n in (person["name"], person["id"], "everyone", person.get("department")) for n in names)


@_read("calendar")
def calendar_list_events(start_date: str = "", end_date: str = "", person: str = "me", query: str = "") -> dict:
    """List meetings and deadlines between two dates (YYYY-MM-DD). person is "me" (default), "all" or a name."""
    start_date = start_date or store.today()
    end_date = end_date or start_date[:8] + "31"
    items = [e for e in store.events() if start_date <= e["start"][:10] <= end_date]
    if person != "all":
        who = store.persona() if person == "me" else store.find_person(person)
        if not who:
            raise ValueError(f"No person called {person}")
        items = [e for e in items if _attends(e, who)]
    if query:
        items = [e for e in items if _matches(e, query)]
    return _all(items)


@_write("calendar")
def calendar_create_event(title: str, start: str, end: str, attendees: list[str] | None = None,
                          location: str = "", notes: str = "") -> dict:
    """Put a meeting in the calendar, organised by you. start and end are "YYYY-MM-DD HH:MM"."""
    event = store.create_event({"title": title, "start": start, "end": end, "attendees": attendees or [],
                                "location": location, "notes": notes})
    return _linked(event, "calendar", event["id"])


# ---------- tickets ----------

def _ticket_summary(t: dict) -> dict:
    return {k: t.get(k) for k in ("id", "title", "status", "priority", "queue", "requester", "assignee", "created")}


@_read("tickets")
def tickets_list(status: str = "", queue: str = "", query: str = "", limit: int = 50) -> dict:
    """List tickets. Filter by status (open, in_progress, waiting, resolved), queue (support, engineering, it, facilities, finance, legal) or words."""
    items = [t for t in store.tickets().values()
             if (not status or t.get("status") == status) and (not queue or t.get("queue") == queue)
             and (not query or _matches(t, query))]
    return _page([_ticket_summary(t) for t in sorted(items, key=lambda t: t["id"], reverse=True)], limit)


@_read("tickets")
def tickets_get(ticket_id: str) -> dict:
    """Read one ticket with its description and comments."""
    try:
        return store.tickets()[ticket_id]
    except KeyError:
        raise ValueError(f"No ticket with id {ticket_id}") from None


@_write("tickets")
def tickets_create(title: str, description: str, queue: str = "support", priority: str = "normal",
                   requester: str = "", assignee: str = "", related: str = "") -> dict:
    """Create a ticket. queue is support, engineering, it, facilities, finance or legal. related can name other tickets, threads or bookings."""
    ticket = store.create_ticket({"title": title, "description": description, "queue": queue, "priority": priority,
                                  "requester": requester, "assignee": assignee, "related": related})
    return _linked(ticket, "tickets", ticket["id"])


@_write("tickets")
def tickets_update(ticket_id: str, status: str = "", assignee: str = "", priority: str = "",
                   queue: str = "", comment: str = "") -> dict:
    """Change a ticket's status, assignee, priority or queue, and optionally add a comment. Empty fields stay as they are."""
    ticket = store.update_record("tickets", ticket_id,
                                 {"status": status, "assignee": assignee, "priority": priority, "queue": queue}, comment)
    return _linked(ticket, "tickets", ticket["id"])


# ---------- CRM ----------

@_read("crm")
def crm_list_accounts(query: str = "", limit: int = 50) -> dict:
    """List corporate accounts in the CRM."""
    return _page([a for a in store.accounts().values() if not query or _matches(a, query)], limit)


@_read("crm")
def crm_list_deals(stage: str = "", owner: str = "", account_id: str = "", query: str = "", limit: int = 50) -> dict:
    """List deals. stage is one of prospect, proposal, negotiation, won, lost. Returns summaries; use crm_get_deal for notes."""
    items = [d for d in store.deals().values()
             if (not stage or d.get("stage") == stage) and (not owner or owner.lower() in d.get("owner", "").lower())
             and (not account_id or d.get("account_id") == account_id) and (not query or _matches(d, query))]
    return _page([{k: v for k, v in d.items() if k != "notes"} for d in items], limit)


@_read("crm")
def crm_get_deal(deal_id: str) -> dict:
    """Read one deal with its call notes, and the account it belongs to."""
    deal = store.deals().get(deal_id)
    if not deal:
        raise ValueError(f"No deal with id {deal_id}")
    return {**deal, "account": store.accounts().get(deal.get("account_id", ""))}


@_write("crm")
def crm_update_deal(deal_id: str, stage: str = "", amount: int = 0, close_date: str = "",
                    loss_reason: str = "", next_step: str = "", note: str = "") -> dict:
    """Update a deal's stage, amount, close date, loss reason or next step, and optionally add a note. Empty fields stay as they are."""
    deal = store.update_record("deals", deal_id, {"stage": stage, "amount": amount or None, "close_date": close_date,
                                                  "loss_reason": loss_reason, "next_step": next_step}, note)
    return _linked(deal, "crm", "deal", deal["id"])


@_write("crm")
def crm_add_account_note(account_id: str, note: str) -> dict:
    """Add a note to a corporate account."""
    account = store.update_record("accounts", account_id, {}, note)
    return _linked(account, "crm", "account", account["id"])


# ---------- drive ----------

@_read("drive")
def drive_list(folder: str = "") -> dict:
    """List files in the shared drive, optionally under one folder. To work on a file itself, fetch a copy with drive_download."""
    folder = folder.strip("/")
    return _all([_public(f) for f in store.drive_files() if not folder or f["path"].startswith(folder + "/")])


@_read("drive")
def drive_read(path: str) -> dict:
    """Read a drive file as text. Spreadsheets, documents and decks are converted to text. For the file itself, for example a deck template to build on, use drive_download."""
    return _public(store.drive_text(path))


@_read("drive")
def drive_search(query: str) -> dict:
    """Find drive files whose name or text contains every word of the query."""
    hits = []
    for f in store.drive_files():
        if _matches(f["path"], query) or _matches(store.drive_text(f["path"])["text"], query):
            hits.append(_public(f))
    return _all(hits)


@_read("drive")
def drive_download(path: str, to_folder: str) -> dict:
    """Copy a drive file in its original format (.pptx, .xlsx, .docx, ...) into a folder on this computer, such as your working folder. Returns saved_to, the path of the copy. Work on the copy; the drive itself only changes through drive_upload."""
    return store.drive_download(path, to_folder)


@_write("drive")
def drive_upload(local_file: str, path: str) -> dict:
    """Put a file from this computer on the shared drive, for example a finished deck as sales/q3-review.pptx. local_file is the absolute path of the file; path is where it goes in the drive."""
    saved = store.drive_upload(local_file, path)
    return _linked(saved, "drive", saved["path"])


@_write("drive")
def drive_write_text(path: str, content: str) -> dict:
    """Save a text file (.md, .csv, .txt) to the shared drive, for example a memo under finance/ or a report under reports/."""
    saved = _public(store.drive_write(path, content))
    return _linked(saved, "drive", saved["path"])


# ---------- bookings snapshot (read only) ----------

def _route_lookup() -> dict[int, dict]:
    return {f["flight_id"]: f for f in store.snapshot("flights")}


def _in_range(value: str, start: str, end: str) -> bool:
    return (not start or value[:10] >= start) and (not end or value[:10] <= end)


@_read("bookings")
def bookings_list_flights(origin: str = "", destination: str = "") -> dict:
    """List flights with route, departure, base price and the price of each seat class."""
    return _all([f for f in store.snapshot("flights")
                 if (not origin or f["origin"].lower() == origin.lower())
                 and (not destination or f["destination"].lower() == destination.lower())])


@_read("bookings")
def bookings_find_customers(query: str = "") -> dict:
    """Find travellers by name or mail address. Returns the user_id used in bookings and quotes."""
    return _all([u for u in store.snapshot("users") if not query or _matches(u, query)])


@_read("bookings")
def bookings_search(user_id: int = 0, flight_id: int = 0, status: str = "", seat_class: str = "",
                    booked_from: str = "", booked_to: str = "", limit: int = 50, offset: int = 0) -> dict:
    """Search bookings. Dates are YYYY-MM-DD on booking_time. For totals use bookings_summary; for all rows read the drive exports."""
    flights = _route_lookup()
    rows = [dict(b, route=f"{flights[b['flight_id']]['origin']} to {flights[b['flight_id']]['destination']}")
            for b in store.snapshot("bookings")
            if (not user_id or b["user_id"] == user_id) and (not flight_id or b["flight_id"] == flight_id)
            and (not status or b["status"] == status) and (not seat_class or b["seat_class"] == seat_class)
            and _in_range(b["booking_time"], booked_from, booked_to)]
    return _page(rows, limit, offset)


@_read("bookings")
def bookings_summary(group_by: str = "route", booked_from: str = "", booked_to: str = "") -> dict:
    """Count bookings and sum price_paid, grouped by route, seat_class, status, month or week (ISO week of booking_time)."""
    from datetime import date
    flights = _route_lookup()

    def key(b: dict) -> str:
        if group_by == "route":
            f = flights[b["flight_id"]]
            return f"{f['origin']} to {f['destination']}"
        if group_by == "month":
            return b["booking_time"][:7]
        if group_by == "week":
            y, w, _ = date.fromisoformat(b["booking_time"][:10]).isocalendar()
            return f"{y}-W{w:02d}"
        if group_by in ("seat_class", "status"):
            return b[group_by]
        raise ValueError("group_by must be route, seat_class, status, month or week")

    groups: dict[str, dict] = {}
    for b in store.snapshot("bookings"):
        if not _in_range(b["booking_time"], booked_from, booked_to):
            continue
        g = groups.setdefault(key(b), {group_by: key(b), "bookings": 0, "cancelled": 0, "price_paid_total": 0,
                                       "price_paid_excluding_cancelled": 0})
        g["bookings"] += 1
        g["price_paid_total"] += b["price_paid"]
        if b["status"] == "cancelled":
            g["cancelled"] += 1
        else:
            g["price_paid_excluding_cancelled"] += b["price_paid"]
    return _all(sorted(groups.values(), key=lambda g: g[group_by]))


@_read("bookings")
def bookings_search_quotes(user_id: int = 0, flight_id: int = 0, seat_class: str = "", status: str = "",
                           booking_id: int = 0, limit: int = 50, offset: int = 0) -> dict:
    """Search price quotes from the hold service. status is converted or expired; converted quotes carry a booking_id."""
    rows = [q for q in store.snapshot("quotes")
            if (not user_id or q["user_id"] == user_id) and (not flight_id or q["flight_id"] == flight_id)
            and (not seat_class or q["seat_class"] == seat_class) and (not status or q["status"] == status)
            and (not booking_id or q["booking_id"] == booking_id)]
    return _page(rows, limit, offset)


@_read("bookings")
def bookings_quote_mismatches(user_id: int = 0, flight_id: int = 0, seat_class: str = "", mismatch_only: bool = True,
                              limit: int = 50, offset: int = 0) -> dict:
    """Join converted quotes to their bookings and compare quoted_price with price_paid. difference is price_paid minus quoted_price: positive means the customer was charged more than quoted. summary covers every pair matching the filters; items are the pairs (only those with a difference when mismatch_only is true), largest overcharge first."""
    flights = _route_lookup()
    bookings = {b["booking_id"]: b for b in store.snapshot("bookings")}
    pairs = []
    for q in store.snapshot("quotes"):
        b = bookings.get(q["booking_id"]) if q["status"] == "converted" else None
        if not b:
            continue
        if ((user_id and b["user_id"] != user_id) or (flight_id and b["flight_id"] != flight_id)
                or (seat_class and b["seat_class"] != seat_class)):
            continue
        f = flights[b["flight_id"]]
        pairs.append({"booking_id": b["booking_id"], "quote_id": q["quote_id"], "user_id": b["user_id"],
                      "flight_id": b["flight_id"], "route": f"{f['origin']} to {f['destination']}",
                      "seat_class": b["seat_class"], "booking_status": b["status"],
                      "quoted_price": q["quoted_price"], "price_paid": b["price_paid"],
                      "difference": b["price_paid"] - q["quoted_price"]})
    more = [p for p in pairs if p["difference"] > 0]
    less = [p for p in pairs if p["difference"] < 0]
    summary = {
        "pairs": len(pairs),
        "mismatches": len(more) + len(less),
        "matching": len(pairs) - len(more) - len(less),
        "charged_more": {"count": len(more), "total_difference": sum(p["difference"] for p in more)},
        "charged_less": {"count": len(less), "total_difference": sum(p["difference"] for p in less)},
    }
    items = [p for p in pairs if p["difference"]] if mismatch_only else pairs
    items.sort(key=lambda p: (-p["difference"], p["booking_id"]))
    return {"summary": summary, **_page(items, limit, offset)}
