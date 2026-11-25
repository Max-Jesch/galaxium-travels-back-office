"""The routine of a working company, generated from the booking snapshot and the cast.

`scripts/build.py` calls `build()` after the snapshot is written. Everything lands in
`content/generated/`, which is wiped first, so nothing in there is edited by hand:

    chat/       daily booking numbers, deployments, the departure log
    mail/       daily summaries, vendor invoices, customer mail to support, notices
    tickets.yaml
    calendar/   recurring meetings and the dates people talk about
    drive/      registers, exports, weekly reports, costings, decks

It uses its own random generator, so adding to it never changes the snapshot.
"""

from __future__ import annotations

import csv
import random
import shutil
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import yaml

import documents
import office

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "content" / "generated"
COMPANY = ROOT / "content" / "company"
TODAY = date(2099, 11, 4)
FIXED_TIME = datetime(2099, 11, 4, 9, 0)
PACK_BOOKINGS = {265, 309, 335, 338}  # named in the scenario packs: leave them out of invented cases
rng = random.Random(4711)

DAY = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
MONTH = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
         "November", "December"]


def money(n: int) -> str:
    return f"{n:,}"


def nice(d: date) -> str:
    return f"{DAY[d.weekday()]} {d.day} {MONTH[d.month]}"


def days(start: date, end: date):
    while start <= end:
        yield start
        start += timedelta(days=1)


def read(table: str) -> list[dict]:
    with (ROOT / "snapshot" / f"{table}.csv").open(encoding="utf-8", newline="") as f:
        rows = list(csv.DictReader(f))
    for row in rows:
        for key, value in row.items():
            if value.isdigit():
                row[key] = int(value)
    return rows


def write(path: str, text: str) -> None:
    target = OUT / path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(text, encoding="utf-8")


# Who keeps each generated workbook, and when it was last touched (shown in the drive).
FILE_META = {
    "finance/bookings-export-2099-08.xlsx": ("Tomás Rivera", "2099-09-01"), "finance/bookings-export-2099-09.xlsx": ("Tomás Rivera", "2099-10-01"),
    "finance/close/2099-09/revenue-pack-2099-09.xlsx": ("Tomás Rivera", "2099-10-01"), "finance/close/2099-10/deferred-revenue-2099-10.xlsx": ("Tomás Rivera", "2099-11-03"),
    "finance/close/2099-10/close-checklist.xlsx": ("Chidi Eze", "2099-11-03"), "finance/budget-2100/budget-template.xlsx": ("Tomás Rivera", "2099-10-20"),
    "finance/expenses/expense-reports-2099-09.xlsx": ("Hana Kobayashi", "2099-10-09"), "finance/ap/invoice-register-2099-10.xlsx": ("Beatrix Novak", "2099-11-02"),
    "ops/crew/relief-crew-tombaugh-costing.xlsx": ("Andrei Popescu", "2099-10-27"), "ops/crew/venus-second-crew-costing.xlsx": ("Andrei Popescu", "2099-11-03"),
    "ops/fleet/fleet-register.xlsx": ("Zainab Hussein", "2099-10-30"), "ops/schedule/scheduled-service-2099-2100.xlsx": ("Yuki Tanaka", "2099-10-19"),
    "ops/ground/tycho-turnarounds-2099-10.xlsx": ("Lucia Bianchi", "2099-10-30"), "sales/rate-card-2099.xlsx": ("Sinead Byrne", "2099-01-12"),
    "legal/vendor-register.xlsx": ("Sam Whitlock", "2099-10-30"), "people/org/headcount-2099-10.xlsx": ("Dorothea Klein", "2099-10-30"),
    "it/asset-register.xlsx": ("Ling Zhou", "2099-10-26"),
}
STATUS = {"status": {"on hold": office.WARN, "waiting": office.AMBER, "open": office.AMBER, "returned": office.AMBER, "not signed": office.WARN,
                     "refit": office.AMBER, "under review": office.AMBER, "done": office.OK},
          "over_target": {"yes": office.WARN}}


def xlsx(path: str, sheets: dict[str, list[dict]], owner: str | None = None, modified: str | None = None, **options) -> None:
    """A workbook in the drive. An export pulled from a system looks like one; everything else is in the house style."""
    known = FILE_META.get(path, (None, None))
    every = {name: {"status": STATUS, **options.get(name, {})} for name in sheets}
    office.table_book(OUT / "drive" / path, sheets, plain=path.startswith("finance/bookings-export"),
                      owner=owner or known[0], modified=modified or known[1], **every)


def thread(slug: str, subject: str, labels: list[str], messages: list[tuple[str, str, str, str]], cc: str = "") -> None:
    """messages: (from, to, 'YYYY-MM-DD HH:MM', body)."""
    head = yaml.safe_dump({"subject": subject, "labels": labels}, allow_unicode=True, sort_keys=False).strip()
    blocks = [f"From: {frm}\nTo: {to}\n{'Cc: ' + cc + chr(10) if cc else ''}Date: {when}\n\n{body.strip()}" for frm, to, when, body in messages]
    write(f"mail/{slug}.md", f"---\n{head}\n---\n" + "\n\n---\n\n".join(blocks) + "\n")


# ------------------------------------------------------------------ data

class Data:
    def __init__(self) -> None:
        self.users = {u["user_id"]: u for u in read("users")}
        self.flights = {f["flight_id"]: f for f in read("flights")}
        self.bookings = read("bookings")
        self.quotes = read("quotes")
        self.people = yaml.safe_load((COMPANY / "people.yaml").read_text(encoding="utf-8"))
        self.by_day: dict[str, list[dict]] = defaultdict(list)
        for b in self.bookings:
            self.by_day[b["booking_time"][:10]].append(b)
        self.quotes_by_day: dict[str, list[dict]] = defaultdict(list)
        for q in self.quotes:
            self.quotes_by_day[q["created_at"][:10]].append(q)

    def route(self, flight_id: int) -> str:
        f = self.flights[flight_id]
        return f"{f['origin']} to {f['destination']}"

    def gx(self, flight_id: int) -> str:
        return f"GX-{flight_id:03d}"

    def dept(self, name: str) -> list[dict]:
        return [p for p in self.people if p["department"] == name]

    def addr(self, person: dict) -> str:
        return f"{person['name']} <{person['email']}>"


# ------------------------------------------------------------------ chat

def build_chat(d: Data) -> None:
    lines = []
    for day in days(date(2099, 8, 11), TODAY):
        y = day - timedelta(days=1)
        rows = d.by_day.get(str(y), [])
        if not rows:
            text = f"{nice(y)}: no new bookings on the scheduled service."
        else:
            top = Counter(d.route(b["flight_id"]) for b in rows).most_common(1)[0]
            text = (f"{nice(y)}: **{len(rows)}** new bookings, ${money(sum(b['price_paid'] for b in rows))} charged. "
                    f"Busiest route: {top[0]} ({top[1]}).")
            quotes = d.quotes_by_day.get(str(y), [])
            if quotes:
                text += f" Quotes issued: {len(quotes)}, of which {sum(q['status'] == 'converted' for q in quotes)} became bookings."
        lines += [f"## {day}", f"[07:00] Bookings Bot: {text}"]
    write("chat/bookings-feed/daily.md", "\n".join(lines) + "\n")

    lines = []
    for f in d.flights.values():
        dep = datetime.fromisoformat(f["departure_time"].replace("Z", ""))
        arr = datetime.fromisoformat(f["arrival_time"].replace("Z", ""))
        if dep.date() >= TODAY:
            continue
        late = rng.choice([0, 0, 0, 0, 1, 2, 3, 4, 6, 11])
        left = dep + timedelta(minutes=late)
        lines.append((left, f"{d.gx(f['flight_id'])} {f['origin']} → {f['destination']} departed {left:%H:%M}"
                            f"{'' if late < 10 else f', {late} minutes late (late boarding)'}. Scheduled service."))
        if arr.date() < TODAY:
            lines.append((arr + timedelta(minutes=late // 2), f"{d.gx(f['flight_id'])} arrived at {f['destination']}. Doors open."))
    out, current = [], None
    for when, text in sorted(lines):
        if when.date() != current:
            current = when.date()
            out.append(f"## {current}")
        out.append(f"[{when:%H:%M}] Flight Ops Bot: {text}")
    write("chat/ops-control/departures.md", "\n".join(out) + "\n")

    fixed = {date(2099, 10, 1): [("08:38", "hold-service 1.0.0", "Aditi Sharma", "first production release")],
             date(2099, 10, 6): [("13:52", "booking-api 4.11.0", "Jonas Keller", "confirmation time zones, admin sorting")],
             date(2099, 10, 13): [("16:22", "booking-api 4.11.1", "Jonas Keller", "payment connection pool and timeout")],
             date(2099, 10, 14): [("09:40", "booking-api 4.11.2", "Jonas Keller", "confirmation mail retry")]}
    services = [("web-frontend", "Wren Abbott", ["copy changes on the payment page", "accessibility fixes", "dependency updates",
                                                  "countdown styling", "search results layout"]),
                ("warehouse-jobs", "Felix Braun", ["nightly copy of quotes", "new reporting role", "retry on timeout", "bookings view by route"]),
                ("hold-service", "Aditi Sharma", ["expiry job logging", "metrics for hold duration", "expiry job rework"]),
                ("status-checks", "Emeka Obi", ["alert thresholds", "pool usage dashboard"])]
    version = {"web-frontend": [2, 30, 0], "warehouse-jobs": [0, 4, 0], "hold-service": [1, 0, 0], "status-checks": [0, 9, 0]}
    out = []
    for day in days(date(2099, 9, 1), TODAY - timedelta(days=1)):
        posts = list(fixed.get(day, []))
        if day.weekday() < 5:
            for _ in range(rng.choice([0, 0, 1, 1, 1, 2])):
                name, who, notes = rng.choice(services)
                if name == "hold-service" and day < date(2099, 10, 2):
                    continue
                version[name][2] += 1
                posts.append((f"{rng.randint(9, 17):02d}:{rng.randint(0, 59):02d}", f"{name} {'.'.join(map(str, version[name]))}", who, rng.choice(notes)))
        if posts:
            out.append(f"## {day}")
            out += [f"[{t}] Deploy Bot: Deployed `{what}` to production. By {who}: {note}." for t, what, who, note in sorted(posts)]
    write("chat/eng-deploys/deploys.md", "\n".join(out) + "\n")


# ------------------------------------------------------------------ tickets and customer mail

SUPPORT_CASES = [
    # (title, customer writes, agent answers, comment on the ticket)
    ("Seat hold expired before payment",
     "I was booking {route} for {dep} and the page said my hold had expired when I came back with my card. Is my seat gone?",
     "A seat hold lasts 15 minutes. Yours ran out, so the seat went back on sale and nothing was charged. The seat is still free: please ask for a new price and you can book straight away.",
     "Explained the 15 minute hold. Customer booked on the second try."),
    ("Booking rejected, name does not match",
     "Your site refuses my booking for {route} on {dep}. It says the name does not match. I typed my name the way I always do.",
     "The traveller name has to be exactly the name on your account. Your account is under \"{name}\". Please book again with that name, and tell me if the account name itself needs correcting.",
     "Name on account differs from the name entered. Customer rebooked."),
    ("Copy of the receipt for booking {booking}",
     "Could you send me a receipt for my trip {route} on {dep}? I need it for my employer. Booking number {booking}.",
     "Of course. The receipt for booking {booking} is attached to this mail. It shows the amount charged and the seat class.",
     "Receipt sent."),
    ("Baggage allowance in {cls} class",
     "How much baggage can I take on {route}, {dep}, in {cls} class? I am bringing equipment.",
     "In {cls} class you may check in two bags of up to 23 kg each and carry one cabin bag. Equipment above that travels as cargo, which I can arrange if you send me dimensions and weight.",
     "Sent the allowance. Customer will send dimensions if needed."),
    ("Special meal request",
     "I am vegetarian and travelling on {route} on {dep}, booking {booking}. Can I order a meal in advance?",
     "Yes. I have added a vegetarian meal to booking {booking}. You will see it on your confirmation.",
     "Meal preference added to the booking."),
    ("Did my booking go through?",
     "I paid for {route} on {dep} but have not received a confirmation mail. Did it work?",
     "It did. Booking {booking} is confirmed in {cls} class. I have sent the confirmation again; please check your spam folder too.",
     "Booking present. Confirmation resent."),
    ("Travelling with a child",
     "My daughter is nine. Can she travel with me on {route} on {dep}, and does she need her own account?",
     "She is welcome. Every traveller needs a seat and a booking in their own name, and for a child under twelve the booking is made from the parent's account by our team. Reply with her full name and I will add her if a seat is free.",
     "Explained the rule for children. Waiting for the name."),
    ("Medical equipment on board",
     "I use a portable oxygen concentrator. Is that allowed on {route}? I fly on {dep}.",
     "Yes, with a note from your doctor and the device's model number at least 72 hours before departure. I have added the request to booking {booking} and informed the cabin crew.",
     "Medical equipment noted on the booking, crew informed."),
    ("Change of date for booking {booking}",
     "I need to travel a month later than planned. Can you move my booking {booking} ({route}, {dep})?",
     "Bookings cannot be moved. You cancel and book the new date, and the refund follows our refund policy for the days left before departure. I can tell you the exact amounts before you decide.",
     "Explained cancel and rebook. Customer will think about it."),
    ("How early do I need to be at the port?",
     "First time flying with you. When should I be at the port for {route} on {dep}?",
     "Please arrive two hours before departure for inner routes and three hours for Jupiter and beyond. Boarding closes 40 minutes before departure.",
     "Answered."),
]
IT_CASES = [
    ("Laptop will not wake from sleep", "The laptop needs a hard restart every morning.", "Firmware updated. Monitoring for a week."),
    ("Access to the finance folder in the drive", "New joiner needs read access to finance/.", "Access granted after approval from the folder owner."),
    ("VPN certificate not trusted", "After the renewal on 22 October the VPN client refuses to connect.", "Client restarted, clock corrected. Works."),
    ("Second factor: new phone", "Changed phone, the authenticator app is gone.", "Identity checked in person, device re-enrolled."),
    ("Meeting room screen has no signal", "Room Phobos, the screen stays black.", "Cable was unplugged. Fixed and taped."),
    ("Shared mailbox on the phone", "Cannot see the shared mailbox in the phone's mail app.", "Added through the account settings. Guide sent."),
    ("Printer on the second floor", "Prints blank pages.", "Toner replaced."),
    ("Request for a second monitor", "Would like a second monitor for close week.", "Delivered from stock."),
    ("Account locked", "Three wrong passwords after the holiday.", "Unlocked after identity check."),
    ("Read access to the warehouse", "Needs the reporting role for the nightly copies.", "Granted after approval from Engineering."),
]
ENG_CASES = [
    ("Confirmation mail shows the wrong time zone", "Departure time on confirmations is in UTC, customers expect station time.", "Fixed in 4.11."),
    ("Admin list of bookings is not sorted", "Support cannot find the latest bookings quickly.", "Sorted by booking time in 4.11."),
    ("Two confirmation mails for one booking", "After the payment timeouts on 13 October some customers got two mails.", "Fixed in 4.11.2. One booking, one charge in every case."),
    ("Search shows a full flight as available", "A flight with ten bookings still shows a seat for a few seconds.", "Cache shortened to five seconds."),
    ("Hold expires at once on a slow connection", "Timer starts when the quote is created on the server.", "Issue 55, planned for 4.12."),
    ("Export of bookings times out for a whole year", "The admin export fails for long ranges.", "Use the warehouse for long ranges. Documented."),
]
FACILITY_CASES = [
    ("Heating on the third floor", "Too cold at the window desks.", "Valve replaced."),
    ("Badge does not open the bike room", "Since the lift maintenance.", "Reader reset."),
    ("Desk chair broken", "Gas lift gone.", "Chair replaced."),
    ("Coffee machine on the fourth floor", "Leaks.", "Service visit done."),
]


def build_tickets_and_support_mail(d: Data) -> None:
    support = [p for p in d.dept("Support") if "Agent" in p["role"]]
    it = d.dept("IT")
    staff = [p["name"] for p in d.people]
    eligible = [b for b in d.bookings if b["booking_id"] not in PACK_BOOKINGS and b["status"] != "cancelled"]
    tickets: list[dict] = []
    stamps = sorted(datetime(2099, 8, 10) + timedelta(minutes=rng.randint(8 * 60, 85 * 24 * 60)) for _ in range(197))
    stamps = [s.replace(hour=rng.randint(7, 19)) for s in stamps]
    old = [s for s in stamps if s < datetime(2099, 10, 14)][-150:]
    new = [s for s in stamps if s >= datetime(2099, 10, 14)][:46]
    numbered = [(1000 - len(old) + 1 + i, s) for i, s in enumerate(sorted(old))] + [(1014 + i + (1 if 1014 + i >= 1019 else 0), s) for i, s in enumerate(sorted(new))]
    mails = 0
    for number, created in numbered:
        tid = f"GT-{number:04d}"
        age = (datetime(2099, 11, 4, 9) - created).days
        kind = rng.choices(["support", "it", "engineering", "facilities"], [62, 20, 10, 8])[0]
        status = "resolved" if age > 6 or rng.random() < 0.4 else rng.choice(["open", "in_progress", "waiting"])
        answered = created + timedelta(minutes=rng.randint(12, 240))
        if kind == "support":
            b = rng.choice([x for x in eligible if x["booking_time"][:10] <= str(created.date())] or eligible)
            user, f = d.users[b["user_id"]], d.flights[b["flight_id"]]
            agent = rng.choice(support)
            title, ask, answer, note = rng.choice(SUPPORT_CASES)
            dep = datetime.fromisoformat(f["departure_time"].replace("Z", ""))
            fill = {"route": d.route(b["flight_id"]), "dep": f"{dep.day} {MONTH[dep.month]}", "booking": b["booking_id"],
                    "cls": b["seat_class"], "name": user["name"]}
            if "hold" in title.lower() and created < datetime(2099, 10, 1):
                title, ask, answer, note = SUPPORT_CASES[5]
            ticket = {"id": tid, "title": title.format(**fill), "status": status, "priority": "normal", "queue": "support",
                      "requester": user["email"], "assignee": agent["name"], "created": f"{created:%Y-%m-%d %H:%M}",
                      "related": f"booking {b['booking_id']}", "description": ask.format(**fill),
                      "comments": [{"author": agent["name"], "date": f"{answered:%Y-%m-%d %H:%M}", "text": note}] if status != "open" else []}
            if rng.random() < 0.45 and status != "open":
                mails += 1
                first = agent["name"].split()[0]
                thread(f"case-{tid.lower()}", ticket["title"], ["support"], [
                    (f"{user['name']} <{user['email']}>", "Galaxium Support <support@galaxium.example>", f"{created - timedelta(minutes=9):%Y-%m-%d %H:%M}",
                     f"Hello,\n\n{ask.format(**fill)}\n\n{user['name']}"),
                    (d.addr(agent), f"{user['name']} <{user['email']}>", f"{answered:%Y-%m-%d %H:%M}",
                     f"Hello {user['name']},\n\n{answer.format(**fill)}\n\nYour reference is {tid}.\n\n{first}, Galaxium Support"),
                ])
        else:
            cases, queue, owners = {"it": (IT_CASES, "it", [p["name"] for p in it]),
                                    "engineering": (ENG_CASES, "engineering", ["Jonas Keller", "Aditi Sharma", "Wren Abbott", "Emeka Obi"]),
                                    "facilities": (FACILITY_CASES, "facilities", ["Agnes Wójcik"])}[kind]
            title, description, note = rng.choice(cases)
            owner = rng.choice(owners)
            requester = rng.choice([s for s in staff if s != owner])
            if kind == "engineering":
                requester = rng.choice(["Nadia Haddad", "Kofi Mensah", "Thandi Dlamini", "Greta Olsen", "Maya Goldberg"])
            ticket = {"id": tid, "title": title, "status": status, "priority": rng.choice(["low", "normal", "normal", "high"]),
                      "queue": queue, "requester": requester, "assignee": owner, "created": f"{created:%Y-%m-%d %H:%M}",
                      "description": description,
                      "comments": [{"author": owner, "date": f"{answered + timedelta(hours=rng.randint(1, 40)):%Y-%m-%d %H:%M}", "text": note}] if status == "resolved" else []}
        tickets.append(ticket)
    tickets.append({
        "id": "GT-1019", "title": "Assistance requests do not reach the handler at Tycho", "status": "resolved", "priority": "high",
        "queue": "support", "requester": "judy@pluto.com", "assignee": "Anneli Korhonen", "created": "2099-10-19 09:05",
        "description": "Three travellers in October booked wheelchair assistance for arrival on the Moon and nobody was there. Requests were sent to the ground handler on time.",
        "comments": [
            {"author": "Anneli Korhonen", "date": "2099-10-20 13:10", "text": "Escalated to Operations in #support-escalations."},
            {"author": "Lucia Bianchi", "date": "2099-10-20 15:02", "text": "Tycho's assistance mailbox was not read at weekends. Changed as of today."},
            {"author": "Anneli Korhonen", "date": "2099-10-21 15:30", "text": "Customers informed, travel credit of $100,000 added for each."}]})
    tickets.sort(key=lambda t: t["id"])
    write("tickets.yaml", "# Generated by scripts/routine.py. Do not edit by hand.\n" + yaml.safe_dump(tickets, allow_unicode=True, sort_keys=False, width=1000))
    print(f"routine: {len(tickets)} tickets, {mails} customer mail threads")


# ------------------------------------------------------------------ mail

VENDORS = [
    ("Helios Fuels", "ar@heliosfuels.example", "HF", "Fuel delivery, Cape Meridian", (1_200_000, 1_900_000)),
    ("Lagrange Catering", "billing@lagrangecatering.example", "LC", "Cabin catering", (60_000, 140_000)),
    ("Olympus Port Authority", "billing@olympusport.example", "OP", "Port and pad fees, Olympus", (180_000, 320_000)),
    ("Aphrodite Station Operations", "accounts@aphroditestation.example", "AS", "Berth fees, Aphrodite Station", (150_000, 260_000)),
    ("Ferrypay", "invoices@ferrypay.example", "FP", "Payment processing fees", (20_000, 90_000)),
    ("Cape Meridian Facilities", "billing@capemeridianfm.example", "CM", "Office cleaning and maintenance, Earth HQ", (30_000, 70_000)),
    ("Northlight Telecom", "billing@northlight.example", "NT", "Connectivity, three stations", (40_000, 60_000)),
    ("Tombaugh Station Cooperative", "treasury@tombaughcoop.example", "TS", "Berth and servicing, Tombaugh", (400_000, 700_000)),
    ("Outer Belt Port Services", "ar@obports.example", "OB", "Handling, Ganymede Gate", (220_000, 380_000)),
    ("Skyward Uniforms", "orders@skywarduniforms.example", "SU", "Crew uniforms", (15_000, 55_000)),
    ("Quill & Ledger", "office@quillledger.example", "QL", "Office supplies", (4_000, 18_000)),
    ("Meridian Crew Hotels", "groups@meridiancrewhotels.example", "MH", "Crew accommodation", (50_000, 160_000)),
    ("Orbital Parts Supply", "sales@orbitalparts.example", "OS", "Spare parts", (80_000, 450_000)),
    ("Brightwell Training", "admin@brightwelltraining.example", "BT", "Simulator sessions", (90_000, 210_000)),
    ("Pulse Media", "billing@pulsemedia.example", "PM", "Paid media", (60_000, 240_000)),
]


def build_mail(d: Data) -> None:
    finance = "Finance <finance@galaxium.example>"
    system = "Booking System <noreply@galaxium.example>"
    for day in days(date(2099, 10, 1), TODAY):
        y = day - timedelta(days=1)
        rows, quotes = d.by_day.get(str(y), []), d.quotes_by_day.get(str(y), [])
        by_class = Counter(b["seat_class"] for b in rows)
        body = (f"Daily summary for {nice(y)} {y.year} (scheduled service, by booking time, UTC)\n\n"
                f"New bookings: {len(rows)}\n"
                f"  economy {by_class['economy']}, business {by_class['business']}, galaxium {by_class['galaxium']}\n"
                f"Amount charged: ${money(sum(b['price_paid'] for b in rows))}\n"
                + (f"Quotes issued: {len(quotes)} (became bookings: {sum(q['status'] == 'converted' for q in quotes)})\n" if y >= date(2099, 10, 1) else "")
                + "\nThis mail is sent automatically at 07:00 UTC. The figures are a count at the time of sending; "
                  "the booking system is the record.")
        thread(f"daily-summary-{y}", f"Booking system: daily summary for {y}", ["automatic", "bookings"],
               [(system, f"{finance}, Operations Control <opscontrol@galaxium.example>", f"{day} 07:00", body)])

    # Vendor invoices of October: a register for the drive and a mail for some of them.
    register, number = [], 3100
    for i in range(209):
        name, sender, prefix, what, (low, high) = VENDORS[i % len(VENDORS)] if i < 30 else rng.choice(VENDORS)
        received = date(2099, 10, 1) + timedelta(days=rng.randint(0, 30))
        number += rng.randint(1, 9)
        amount = rng.randrange(low, high, 500)
        register.append({"received": str(received), "vendor": name, "invoice": f"{prefix}-{number}", "description": what,
                         "amount": amount, "status": "paid" if received < date(2099, 10, 24) else "approved", "note": "", "_from": sender})
    register += [
        {"received": "2099-10-05", "vendor": "Helios Fuels", "invoice": "HF-22871", "description": "Fuel delivery 28 September, delivery note 77120",
         "amount": 1840000, "status": "paid", "note": "", "_from": ""},
        {"received": "2099-10-12", "vendor": "Helios Fuels", "invoice": "HF-22917", "description": "Fuel delivery 28 September, delivery note 77120",
         "amount": 1840000, "status": "on hold", "note": "Duplicate of HF-22871. Credit note HF-CN-0388 received 19 October.", "_from": ""},
        {"received": "2099-10-27", "vendor": "Lagrange Catering", "invoice": "LC-4471", "description": "Cabin catering, week 43",
         "amount": 96500, "status": "on hold", "note": "Quantity dispute: 140 meals invoiced, 112 delivered.", "_from": ""},
    ]
    register.sort(key=lambda r: (r["received"], r["invoice"]))
    assert len(register) == 212
    xlsx("finance/ap/invoice-register-2099-10.xlsx", {"Invoices received October": [{k: v for k, v in r.items() if k != "_from"} for r in register]})
    for r in [r for r in register if r["_from"]][::7]:
        thread(f"invoice-{r['invoice'].lower()}", f"Invoice {r['invoice']}: {r['description']}", ["payables", "automatic"],
               [(f"{r['vendor']} <{r['_from']}>", "Invoices <invoices@galaxium.example>", f"{r['received']} {rng.randint(6, 17):02d}:{rng.randint(0, 59):02d}",
                 f"Dear customer,\n\nPlease find attached invoice {r['invoice']} for {r['description'].lower()}.\n\n"
                 f"Amount due: ${money(r['amount'])}. Payment terms as agreed.\n\n{r['vendor']}")])

    everyone = "All Staff <all@galaxium.example>"
    menus = ["Europa noodle soup, roast vegetables, lemon tart", "Lentil curry, grilled fish, rice pudding", "Mushroom risotto, chicken with herbs, fruit",
             "Tomato soup, vegetable lasagne, apple cake", "Chickpea stew, baked salmon, yoghurt with honey", "Pumpkin soup, pasta with pesto, chocolate mousse"]
    for i, monday in enumerate(date(2099, 9, 28) + timedelta(days=7 * k) for k in range(6)):
        thread(f"canteen-menu-{monday}", f"Canteen menu, week of {monday.day} {MONTH[monday.month]}", ["all-staff", "canteen"],
               [("Agnes Wójcik <agnes.wojcik@galaxium.example>", everyone, f"{monday} 08:15",
                 f"Good morning,\n\nThis week in the Earth HQ canteen:\n\n" + "\n".join(f"- {DAY[k]}: {menus[(i + k) % len(menus)]}" for k in range(5))
                 + "\n\nLuna and Mars colleagues: the lunch allowance applies as usual.\n\nAgnes")])
    for when, month in [("2099-09-25 09:00", "September"), ("2099-10-23 09:00", "October")]:
        thread(f"payslip-{month.lower()}", f"Your payslip for {month} is available", ["all-staff", "payroll", "automatic"],
               [("Payroll <payroll@galaxium.example>", everyone, when,
                 f"Your payslip for {month} 2099 is available in the payroll system. Pay day is the 28th.\n\n"
                 "Questions about your payslip go to Hana Kobayashi. Please do not reply to this mail.")])
    for when, what in [("2099-10-09 16:00", "the drive will be read-only on Saturday 10 October between 06:00 and 08:00 UTC for a storage upgrade"),
                       ("2099-10-21 16:00", "VPN certificates are renewed tomorrow, 22 October. You will be asked to confirm the new certificate once"),
                       ("2099-10-30 16:00", "mail will be unavailable for about ten minutes on Sunday 1 November at 05:00 UTC")]:
        thread(f"it-maintenance-{when[:10]}", f"Planned maintenance: {what.split(' will ')[0].split(' are ')[0]}", ["it", "all-staff"],
               [("IT <it@galaxium.example>", everyone, when, f"Hello,\n\nA short notice: {what}.\n\nNothing for you to do. Questions in #help-it.\n\nBruno and Ling")])


# ------------------------------------------------------------------ calendar

def build_calendar(d: Data) -> None:
    events: list[dict] = []

    def add(title: str, day: date, start: str = "", end: str = "", organizer: str = "", attendees: list[str] | None = None,
            location: str = "", notes: str = "", repeats: str = "") -> None:
        events.append({"id": f"EV-{len(events) + 1:04d}", "title": title,
                       "start": f"{day} {start}".strip(), "end": f"{day} {end}".strip() if end else "",
                       "organizer": organizer, "attendees": attendees or [], "location": location,
                       **({"notes": notes} if notes else {}), **({"repeats": repeats} if repeats else {})})

    first, last = date(2099, 9, 28), date(2099, 11, 27)
    for day in days(first, last):
        wd = day.weekday()
        if wd > 4:
            continue
        add("Engineering stand-up", day, "09:30", "09:45", "Marcus Lindqvist", ["Engineering"], "Video", repeats="Every working day")
        add("Operations control handover", day, "08:00", "08:15", "Sofia Lindgren", ["Hiro Matsuda", "Priya Raman"], "Operations control room", repeats="Every working day")
        if wd == 0:
            if day != date(2099, 11, 2):
                add("Route review", day, "09:00", "10:00", "Priya Raman", ["Yuki Tanaka", "Olu Adeyemi", "Tomás Rivera", "Amara Okafor"], "Room Phobos",
                    "Notes go to ops/route-reviews in the drive. One page per open item, sent the Friday before.", "Every Monday")
            add("Support: last week's numbers", day, "09:00", "09:30", "Nadia Haddad", ["Support"], "Room Deimos", repeats="Every Monday")
            add("Legal weekly", day, "11:00", "11:30", "Helena Vogt", ["Legal"], "Helena's office", repeats="Every Monday")
            add("1:1 Helena and Sam", day, "14:00", "14:30", "Helena Vogt", ["Sam Whitlock"], repeats="Every Monday")
        if wd == 1:
            add("Pipeline review", day, "10:00", "10:45", "Amara Okafor", ["Sales"], "Room Ceres", "Read from the CRM. Update your deals by Monday evening.", "Every Tuesday")
            add("1:1 Priya and Yuki", day, "14:00", "14:30", "Priya Raman", ["Yuki Tanaka"], repeats="Every Tuesday")
            add("Marketing weekly", day, "11:00", "11:45", "Elise Fontaine", ["Marketing"], "Room Vesta", repeats="Every Tuesday")
        if wd == 2:
            add("Finance team meeting", day, "11:00", "11:45", "Ingrid Solberg", ["Finance"], "Room Ceres", repeats="Every Wednesday")
            add("1:1 Amara and Dev", day, "15:00", "15:30", "Amara Okafor", ["Dev Malhotra"], repeats="Every Wednesday")
            add("1:1 Amara and Lena", day, "16:00", "16:30", "Amara Okafor", ["Lena Fischer"], "Video", repeats="Every Wednesday")
            add("Product and Engineering planning", day, "13:00", "14:00", "Maya Goldberg", ["Marcus Lindqvist", "Aditi Sharma", "Jonas Keller", "Wren Abbott", "Clara Jensen"], "Room Vesta", repeats="Every Wednesday")
        if wd == 3:
            add("1:1 Ingrid and Tomás", day, "14:00", "14:30", "Ingrid Solberg", ["Tomás Rivera"], "Ingrid's office", repeats="Every Thursday")
            add("Five-a-side", day, "18:30", "19:30", "Chidi Eze", ["Tomás Rivera", "Emeka Obi", "Jonas Keller", "Kofi Mensah", "Samir Chaudhry"], "Sports hall, Cape Meridian", repeats="Every Thursday")
        if wd == 4:
            add("1:1 Olu and Priya", day, "10:00", "10:30", "Olu Adeyemi", ["Priya Raman"], repeats="Every Friday")
            add("Leadership: week in review", day, "16:00", "16:30", "Renata Castellanos", ["Leadership"], "Video", repeats="Every Friday")

    heads = ["Leadership", "Ingrid Solberg", "Amara Okafor", "Priya Raman", "Helena Vogt", "Elise Fontaine", "Dorothea Klein"]
    D = date
    add("All-hands", D(2099, 10, 1), "15:00", "16:00", "Renata Castellanos", ["everyone"], "Atrium, Earth HQ, and video")
    add("All-hands", D(2099, 11, 5), "15:00", "16:00", "Renata Castellanos", ["everyone"], "Atrium, Earth HQ, and video",
        "1. October results (Ingrid)\n2. Winter on Mars (Elise)\n3. After the hold service (Farid)\n4. Questions")
    add("Leadership meeting", D(2099, 10, 15), "10:00", "12:00", "Renata Castellanos", heads, "Boardroom")
    add("Leadership meeting", D(2099, 11, 12), "10:00", "12:00", "Renata Castellanos", heads, "Boardroom",
        "1. October results (Ingrid)\n2. Win-loss, third quarter (Amara)\n3. Fleet options 2100, first look (Priya, Olu)\n4. After the hold service (Farid)\n\nBoard one-pagers are due today.")
    add("Board meeting", D(2099, 11, 19), "09:00", "13:00", "Renata Castellanos", ["Leadership", "Ingrid Solberg", "Helena Vogt"], "Boardroom")
    add("Postmortem: payment timeouts", D(2099, 10, 16), "11:00", "12:00", "Emeka Obi", ["Engineering", "Nadia Haddad", "Sofia Lindgren"], "Room Vesta")
    for n, day in enumerate(days(D(2099, 11, 2), D(2099, 11, 6)), 1):
        add(f"October close, day {n}", day, organizer="Chidi Eze", attendees=["Finance"], notes="Calendar: finance/close/close-calendar-2099-10.md")
    add("Expense reports for October due 18:00", D(2099, 11, 3), organizer="Hana Kobayashi", attendees=["everyone"])
    add("Benefits drop-in", D(2099, 11, 3), "12:30", "13:15", "Dorothea Klein", ["everyone"], "Room Deimos")
    add("Benefits drop-in", D(2099, 11, 10), "12:30", "13:15", "Dorothea Klein", ["everyone"], "Room Deimos")
    add("Venus planning workshop", D(2099, 11, 4), "09:00", "17:00", "Yuki Tanaka", ["Andrei Popescu", "Lucia Bianchi", "Zainab Hussein"], "Room Phobos")
    add("Call: Nimbus Freight renewal", D(2099, 11, 5), "11:00", "11:45", "Jamal Carter", ["Amara Okafor", "Dev Malhotra"], "Video", "With Peter Lindholm. Account plan is in sales/accounts in the drive.")
    add("Customer newsletter goes out", D(2099, 11, 5), organizer="Noor Rahimi", attendees=["Marketing"])
    add("Code freeze for release 4.12", D(2099, 11, 6), organizer="Marcus Lindqvist", attendees=["Engineering"])
    add("Quote due: Tharsis Engineering", D(2099, 11, 6), organizer="Bastien Moreau", attendees=["Amara Okafor"])
    add("Budget actuals to departments", D(2099, 11, 6), organizer="Tomás Rivera", attendees=["Finance"])
    add("Multi-factor sign-in becomes mandatory", D(2099, 11, 9), organizer="Bruno Ferreira", attendees=["everyone"])
    add("Release 4.12", D(2099, 11, 10), "07:00", "08:00", "Marcus Lindqvist", ["Engineering", "Greta Olsen"], "Video")
    add("Winter on Mars launch", D(2099, 11, 10), "09:00", "09:30", "Elise Fontaine", ["Marketing", "Amara Okafor", "Nadia Haddad"], "Room Vesta")
    add("Flu vaccinations, 09:00 to 13:00", D(2099, 11, 10), organizer="Agnes Wójcik", attendees=["everyone"], location="Room Deimos")
    add("Final interview: Senior Engineer, candidate A", D(2099, 11, 10), "14:00", "16:30", "Samir Chaudhry", ["Marcus Lindqvist", "Aditi Sharma", "Maya Goldberg"], "Room Ceres")
    add("Final interview: Senior Engineer, candidate B", D(2099, 11, 11), "09:30", "12:00", "Samir Chaudhry", ["Marcus Lindqvist", "Aditi Sharma", "Maya Goldberg"], "Room Ceres")
    add("Security and privacy training due", D(2099, 11, 13), organizer="Isaac Brandt", attendees=["everyone"])
    add("Benefits enrolment closes", D(2099, 11, 13), organizer="Dorothea Klein", attendees=["everyone"])
    add("Performance cycle opens: self reviews", D(2099, 11, 16), organizer="Dorothea Klein", attendees=["everyone"])
    add("Quarterly safety review", D(2099, 11, 17), "10:00", "12:00", "Kwame Asante", ["Priya Raman", "Olu Adeyemi", "Zainab Hussein", "Lucia Bianchi", "Andrei Popescu"], "Room Phobos")
    add("Book club", D(2099, 11, 18), "17:30", "18:30", "Noor Rahimi", ["Viktor Hale", "Hana Kobayashi", "Clara Jensen"], "Canteen")
    add("Budget 2100: first drafts due", D(2099, 11, 20), organizer="Ingrid Solberg", attendees=heads + ["Marcus Lindqvist", "Nadia Haddad", "Bruno Ferreira"])
    for day in days(D(2099, 11, 23), D(2099, 11, 27)):
        add("Auditors on site: interim fieldwork", day, organizer="Chidi Eze", attendees=["Finance"], location="Room Ceres")
    for day in days(D(2099, 11, 24), D(2099, 11, 26)):
        add("Interplanetary Travel Expo, Luna", day, organizer="Pilar Navarro", attendees=["Lena Fischer", "Jamal Carter", "Maya Goldberg", "Elise Fontaine"], location="Hall 2, stand 214")
    add("Budget 2100: review meetings", D(2099, 11, 27), "09:00", "17:00", "Ingrid Solberg", heads, "Boardroom")
    events.sort(key=lambda e: (e["start"], e["id"]))
    write("calendar/events.yaml", "# Generated by scripts/routine.py. Do not edit by hand.\n" + yaml.safe_dump(events, allow_unicode=True, sort_keys=False, width=1000))
    print(f"routine: {len(events)} calendar events")


# ------------------------------------------------------------------ drive

def build_drive(d: Data) -> None:
    name = lambda uid: d.users[uid]["name"]  # noqa: E731
    row = lambda b: {**b, "customer": name(b["user_id"]), "route": d.route(b["flight_id"])}  # noqa: E731

    for month, label in [("2099-08", "August"), ("2099-09", "September")]:
        rows = [row(b) for b in d.bookings if b["booking_time"].startswith(month)]
        xlsx(f"finance/bookings-export-{month}.xlsx", {f"Bookings {label} 2099": rows})
    sept = [b for b in d.bookings if b["booking_time"].startswith("2099-09")]

    def split(rows: list[dict], key) -> list[dict]:
        groups: dict[str, list[dict]] = defaultdict(list)
        for b in rows:
            groups[key(b)].append(b)
        return [{"group": g, "bookings": len(v), "cancelled": sum(b["status"] == "cancelled" for b in v),
                 "charged": sum(b["price_paid"] for b in v),
                 "charged_excluding_cancelled": sum(b["price_paid"] for b in v if b["status"] != "cancelled")}
                for g, v in sorted(groups.items())]

    xlsx("finance/close/2099-09/revenue-pack-2099-09.xlsx", {
        "By route": split(sept, lambda b: d.route(b["flight_id"])), "By seat class": split(sept, lambda b: b["seat_class"]),
        "Source": [{"note": "Booking system, bookings with a booking time in September 2099, pulled 1 October 2099 by Tomás Rivera."}]})

    deferred = [{"booking_id": b["booking_id"], "customer": name(b["user_id"]), "route": d.route(b["flight_id"]),
                 "departure": d.flights[b["flight_id"]]["departure_time"][:10], "seat_class": b["seat_class"], "deferred_amount": b["price_paid"]}
                for b in d.bookings if b["status"] != "cancelled" and b["booking_time"][:10] <= "2099-10-31"
                and d.flights[b["flight_id"]]["departure_time"][:10] > "2099-10-31"]
    xlsx("finance/close/2099-10/deferred-revenue-2099-10.xlsx", {"Deferred at 31 October": deferred, "Method": [
        {"note": "Bookings not cancelled, booked on or before 31 October 2099, departing after 31 October 2099, at the amount charged."},
        {"note": "Source: booking system, 3 November 2099. Prepared by Tomás Rivera."}]})

    checklist = [("Bank reconciliations, four accounts", "Chidi Eze", 1, "done"), ("Payables cut-off", "Beatrix Novak", 1, "done"),
                 ("Payroll accrual", "Hana Kobayashi", 1, "done"), ("Expense reports cut-off", "Hana Kobayashi", 2, "done"),
                 ("Deferred revenue schedule", "Tomás Rivera", 2, "done"), ("Intercompany, Luna entity", "Chidi Eze", 2, "done"),
                 ("Revenue: bookings to ledger", "Tomás Rivera", 3, "open"), ("Revenue: quotes to bookings (new)", "Tomás Rivera", 3, "open"),
                 ("Tycho Ground October invoice: hold or accrue", "Beatrix Novak", 3, "waiting for Legal"),
                 ("Flux commentary", "Chidi Eze", 4, "open"), ("Review", "Ingrid Solberg", 4, "open"), ("Sign-off", "Ingrid Solberg", 5, "open")]
    xlsx("finance/close/2099-10/close-checklist.xlsx", {"October 2099 close": [{"task": t, "owner": o, "close_day": n, "status": s} for t, o, n, s in checklist]})

    lines = ["Revenue, scheduled service", "Revenue, charter", "Ship leases", "Traffic control charges", "Insurance", "Fuel", "Ground handling and port fees", "Crew", "Maintenance", "Staff, non-crew",
             "Marketing", "Software and IT", "Office", "Travel", "Other"]
    xlsx("finance/budget-2100/budget-template.xlsx", {"Budget 2100": [{"line": line, "actual_2026_to_october": "", **{MONTH[m][:3]: "" for m in range(1, 13)}, "comment": ""} for line in lines],
                                                        "Headcount": [{"role": "", "grade": "", "start_month": "", "monthly_cost": "", "reason": ""}]})

    staff = [p["name"] for p in d.people]
    reports = []
    for i in range(61):
        who = "Tomás Rivera" if i == 30 else rng.choice(staff)
        amount = 42000 if i == 30 else rng.randrange(6000, 180000, 500)
        status = "returned: receipt missing, resubmitted and paid" if i == 30 else ("returned: receipt missing" if i in (7, 19, 48) else "paid")
        reports.append({"report": f"ER-2099-{901 + i:04d}", "employee": who, "month": "September", "amount": amount, "lines": rng.randint(1, 9), "status": status})
    xlsx("finance/expenses/expense-reports-2099-09.xlsx", {"Expense reports September": reports})

    # Operations
    monday = date(2099, 8, 10)
    while monday + timedelta(days=6) <= date(2099, 10, 25):
        week = [b for b in d.bookings if str(monday) <= b["booking_time"][:10] <= str(monday + timedelta(days=6))]
        if week:
            summary = [{"measure": "Bookings", "value": len(week)}, {"measure": "Cancelled since", "value": sum(b["status"] == "cancelled" for b in week)},
                       {"measure": "Charged, excluding cancelled (dollars)", "value": sum(b["price_paid"] for b in week if b["status"] != "cancelled")},
                       {"measure": "Week", "value": f"{monday} to {monday + timedelta(days=6)}, by booking date"},
                       {"measure": "Source", "value": "Booking system, pulled the following Monday by Yuki Tanaka"}]
            xlsx(f"ops/weekly-reports/weekly-bookings-{monday.isocalendar().year}-W{monday.isocalendar().week:02d}.xlsx",
                 {"Summary": summary, "By route": split(week, lambda b: d.route(b["flight_id"])), "By seat class": split(week, lambda b: b["seat_class"])},
                 owner="Yuki Tanaka", modified=str(monday + timedelta(days=7)))
        monday += timedelta(days=7)

    xlsx("ops/crew/relief-crew-tombaugh-costing.xlsx", {"Monthly cost": [
        {"item": "Four crew on a six week rotation, salaries and allowances", "dollars_per_month": 1180000},
        {"item": "Housing at Tombaugh Station, four units", "dollars_per_month": 1240000},
        {"item": "Rotation travel for the relief crew", "dollars_per_month": 220000},
        {"item": "Second set of spares stationed at Tombaugh (financing and storage)", "dollars_per_month": 260000},
        {"item": "Total", "dollars_per_month": 2900000}],
        "Assumptions": [{"note": "Prepared by Andrei Popescu, 27 October 2099. Spares line added by Zainab Hussein."},
                        {"note": "Housing is a quote from the Tombaugh Station Cooperative, valid until 31 December 2099."}]})
    xlsx("ops/crew/venus-second-crew-costing.xlsx", {"Monthly cost": [
        {"item": "Three crew based at Aphrodite Station, salaries and allowances", "dollars_per_month": 960000},
        {"item": "Housing at Aphrodite Station", "dollars_per_month": 590000},
        {"item": "Rotation travel", "dollars_per_month": 150000},
        {"item": "Total", "dollars_per_month": 1700000}],
        "Assumptions": [{"note": "Prepared by Andrei Popescu, 3 November 2099."}]})
    fleet = yaml.safe_load((COMPANY / "fleet.yaml").read_text(encoding="utf-8"))
    class_name = {c["id"]: c["name"].removesuffix(" class") for c in fleet["classes"]}
    xlsx("ops/fleet/fleet-register.xlsx", {"Fleet": [
        {"ship": s["name"], "class": class_name[s["class"]], "seats": fleet["cabin"]["seats"], "certified_to": s["register"]["certified_to"], "routes": s["register"]["routes"],
         "status": s["register"]["status"], "next_check": s["register"]["next_check"]} for s in fleet["ships"]]})
    xlsx("ops/schedule/scheduled-service-2099-2100.xlsx", {"Scheduled service": [
        {"flight": d.gx(f["flight_id"]), "origin": f["origin"], "destination": f["destination"], "departure_utc": f["departure_time"].replace("T", " ")[:16],
         "arrival_utc": f["arrival_time"].replace("T", " ")[:16], "economy_seats": f["economy_seats"], "business_seats": f["business_seats"], "galaxium_seats": f["galaxium_seats"]}
        for f in d.flights.values()], "Changes": [{"note": "From the departure on 13 December 2099 the Earth to Venus service leaves at 08:30. The booking system shows the old time until the change is loaded."}]})
    late = {2: 112, 9: 104, 21: 97}
    turn_days = [2, 3, 5, 7, 9, 12, 14, 16, 19, 21, 23, 26, 28, 30]
    xlsx("ops/ground/tycho-turnarounds-2099-10.xlsx", {"Tycho October 2099": [
        {"date": f"2099-10-{day:02d}", "kind": "Scheduled service" if day == 5 else rng.choice(["Luna shuttle", "Luna shuttle", "Charter"]),
         "turnaround_minutes": late.get(day, rng.randint(61, 88)), "target_minutes": 90, "over_target": "yes" if day in late else "no",
         "cause": "Shared fuelling gantry" if day in late else ""} for day in turn_days],
        "Source": [{"note": "Operations control log. Kept by Lucia Bianchi."}]})

    xlsx("sales/rate-card-2099.xlsx", {"Charter rate card 2099": [
        {"route": r, "seat_economy": e, "seat_business": int(e * 2.5), "whole_ship_10_seats": int(e * 2.5 * 10 * 0.8)}
        for r, e in [("Earth to Mars", 1000000), ("Earth to Moon", 500000), ("Mars to Earth", 950000), ("Venus to Earth", 1200000), ("Jupiter to Europa", 2000000),
                     ("Earth to Venus", 1100000), ("Moon to Mars", 800000), ("Mars to Jupiter", 2500000), ("Europa to Earth", 3000000), ("Earth to Pluto", 5000000)]],
        "Notes": [{"note": "Seat prices equal the public base prices. Whole-ship price is ten business seats less 20%."},
                  {"note": "Discounts follow policies/discount-policy.md."}]})

    xlsx("legal/vendor-register.xlsx", {"Vendors": [
        {"vendor": "Tycho Ground Services", "service": "Ground handling, Luna", "contract_ends": "2099-12-31", "monthly_fee": 380000, "automatic_renewal": "no", "owner": "Lucia Bianchi", "status": "Renewal draft received 30 October, under review"},
        {"vendor": "Helios Fuels", "service": "Fuel supply", "contract_ends": "2100-01-31", "monthly_fee": "by delivery", "automatic_renewal": "no", "owner": "Zainab Hussein", "status": "Renewal draft expected in November"},
        {"vendor": "Lagrange Catering", "service": "Cabin catering", "contract_ends": "2100-10-20", "monthly_fee": "by delivery", "automatic_renewal": "no", "owner": "Lucia Bianchi", "status": "Renewed 21 October"},
        {"vendor": "Ferrypay", "service": "Payment processing", "contract_ends": "2100-06-30", "monthly_fee": "per transaction", "automatic_renewal": "yes, 12 months", "owner": "Marcus Lindqvist", "status": "Active"},
        {"vendor": "Olympus Port Authority", "service": "Port and pad fees, Mars", "contract_ends": "2100-06-30", "monthly_fee": 250000, "automatic_renewal": "no", "owner": "Lucia Bianchi", "status": "Active"},
        {"vendor": "Aphrodite Station Operations", "service": "Berth, Venus", "contract_ends": "2100-03-31", "monthly_fee": 210000, "automatic_renewal": "no", "owner": "Lucia Bianchi", "status": "Active"},
        {"vendor": "Tombaugh Station Cooperative", "service": "Berth and servicing, Pluto", "contract_ends": "2100-08-31", "monthly_fee": 540000, "automatic_renewal": "no", "owner": "Lucia Bianchi", "status": "Active"},
        {"vendor": "Northlight Telecom", "service": "Connectivity", "contract_ends": "2100-02-28", "monthly_fee": 52000, "automatic_renewal": "yes, 12 months", "owner": "Bruno Ferreira", "status": "Active"},
        {"vendor": "Mail and collaboration suite", "service": "Mail, chat, drive", "contract_ends": "2099-12-15", "monthly_fee": 61000, "automatic_renewal": "yes, 12 months", "owner": "Bruno Ferreira", "status": "Renews 15 December"},
        {"vendor": "Pulse Feedback", "service": "Customer surveys", "contract_ends": "", "monthly_fee": 14000, "automatic_renewal": "", "owner": "Nadia Haddad", "status": "Not signed: data processing addendum open"}]})

    xlsx("people/org/headcount-2099-10.xlsx", {"Headcount 31 October": [
        {"name": p["name"], "role": p["role"], "department": p["department"], "location": p.get("location", ""),
         "manager": next((m["name"] for m in d.people if m["id"] == p.get("manager")), "")} for p in d.people],
        "By department": [{"department": k, "people": v} for k, v in sorted(Counter(p["department"] for p in d.people).items())]})
    xlsx("it/asset-register.xlsx", {"Laptops": [
        {"asset": f"LT-{200 + i:04d}", "assigned_to": p["name"], "model": rng.choice(["Model A 14", "Model A 16", "Model B 13"]),
         "purchased": f"{rng.choice([2094, 2095, 2096, 2097, 2098])}-{rng.randint(1, 12):02d}", "note": "Replacement ordered, delivery 5 November" if p["name"] in ("Andrei Popescu", "Hana Kobayashi") else ""}
        for i, p in enumerate(d.people)]})


def build() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    data = Data()
    build_chat(data)
    build_tickets_and_support_mail(data)
    build_mail(data)
    build_calendar(data)
    build_drive(data)
    documents.build(data)
    office.write_index(OUT / "drive")
    print(f"routine: {sum(1 for p in OUT.rglob('*') if p.is_file())} files in content/generated")


if __name__ == "__main__":
    build()
