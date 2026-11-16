"""Rebuild everything that is generated: the booking snapshot, the office files in the
drive, and the ground-truth files of the scenario packs.

The output is committed, so nobody needs to run this to use the back office. Run it
only after changing the rules below or the win-loss deals:

    .venv/bin/python scripts/build.py

It is deterministic: the same script always writes the same numbers.
"""

from __future__ import annotations

import csv
import random
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import yaml
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font
from pptx import Presentation
from pptx.util import Pt

ROOT = Path(__file__).resolve().parent.parent
SNAPSHOT = ROOT / "snapshot"
COMPANY = ROOT / "content" / "company"
PACKS = ROOT / "content" / "packs"
TODAY = datetime(2026, 11, 4)
HOLD_SERVICE_LIVE = datetime(2026, 10, 1)
FIXED_TIME = datetime(2026, 11, 4, 9, 0)
rng = random.Random(2026)


def money(n: int) -> str:
    return f"{n:,}"


def write_csv(name: str, rows: list[dict]) -> None:
    SNAPSHOT.mkdir(exist_ok=True)
    with (SNAPSHOT / f"{name}.csv").open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def write_xlsx(path: Path, sheets: dict[str, list[dict]]) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        ws.append(list(rows[0]))
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for row in rows:
            ws.append(list(row.values()))
        ws.freeze_panes = "A2"
        for i, key in enumerate(rows[0]):
            ws.column_dimensions[chr(65 + i)].width = max(12, len(str(key)) + 4, *(len(str(r[key])) + 2 for r in rows[:50]))
    wb.properties.created = wb.properties.modified = FIXED_TIME
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ---------------------------------------------------------------- snapshot

# The ten routes of the galaxium-travels seed, in seed order, with the same base prices.
ROUTES = [
    ("Earth", "Mars", 9, 8, 1000000, 10), ("Earth", "Moon", 10, 4, 500000, 10), ("Mars", "Earth", 12, 8, 950000, 8),
    ("Venus", "Earth", 8, 10, 1200000, 5), ("Jupiter", "Europa", 15, 4, 2000000, 2), ("Earth", "Venus", 7, 8, 1100000, 6),
    ("Moon", "Mars", 11, 8, 800000, 6), ("Mars", "Jupiter", 13, 10, 2500000, 4), ("Europa", "Earth", 9, 12, 3000000, 2),
    ("Earth", "Pluto", 6, 24, 5000000, 3),
]  # origin, destination, departure hour, duration in hours, base price, popularity weight
MONTHS = [(2026, 9), (2026, 10), (2026, 11), (2026, 12), (2027, 1), (2027, 2)]
CLASS_FACTOR = {"economy": 1, "business": 2.5, "galaxium": 5}
SEATS = {"economy": 6, "business": 3, "galaxium": 1}
JAVA_TABLE = {"economy": 500000, "business": 2500000, "first": 5000000}  # PricingService.java


def charged_price(flight: dict, seat_class: str) -> int:
    """What the Python booking backend charges: base price times the seat class factor."""
    return int(flight["base_price"] * CLASS_FACTOR[seat_class])


def quoted_price(flight_id: int, seat_class: str) -> int:
    """What the Java hold service quotes: its own table, an unknown class falls back to
    500,000, and a factor that depends on the flight id instead of the route."""
    return JAVA_TABLE.get(seat_class, 500000) * (10 + flight_id % 3) // 10


def build_snapshot() -> tuple[list[dict], list[dict], list[dict], list[dict]]:
    first = ["Alice", "Bob", "Charlie", "Diana", "Eve", "Frank", "Grace", "Heidi", "Ivan", "Judy"]
    seed_mail = ["alice@example.com", "bob@example.com", "charlie@galaxium.com", "diana@moonmail.com", "eve@marsmail.com",
                 "frank@venusmail.com", "grace@jupiter.com", "heidi@europa.com", "ivan@asteroidbelt.com", "judy@pluto.com"]
    more = ("Kiran Liam Mona Nils Olga Pavel Quinn Rosa Sven Tara Uma Viktor Wanda Xavier Yara Zane Aiko Bruno Chidi Dalia "
            "Emil Farah Goran Hana Idris Jana Kenji Leila Mateo Noor Omar Petra Rafa Sofia Timo Ulla Vera Wim Xenia Yusuf "
            "Zofia Anders Bianca Caleb Dara Elias Freya Gustav Halima Ines").split()
    domains = ["moonmail.example", "marsmail.example", "venusmail.example", "europa.example", "beltpost.example"]
    users = [{"user_id": i + 1, "name": n, "email": seed_mail[i]} for i, n in enumerate(first)]
    users += [{"user_id": 11 + i, "name": n, "email": f"{n.lower()}@{domains[i % 5]}"} for i, n in enumerate(more)]

    flights = []
    for m, (year, month) in enumerate(MONTHS):
        for r, (origin, dest, hour, hours, base, _) in enumerate(ROUTES):
            dep = datetime(year, month, 3 + r * 2, hour)
            flights.append({
                "flight_id": m * 10 + r + 1, "origin": origin, "destination": dest,
                "departure_time": dep.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "arrival_time": (dep + timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%SZ"),
                "base_price": base, "economy_price": base, "business_price": int(base * 2.5), "galaxium_price": base * 5,
                "economy_seats": 6, "business_seats": 3, "galaxium_seats": 1,
            })
    by_id = {f["flight_id"]: f for f in flights}
    departs = {f["flight_id"]: datetime.strptime(f["departure_time"], "%Y-%m-%dT%H:%M:%SZ") for f in flights}

    # Bookings the stories refer to by name. They are placed first so their seats are free.
    fixed = [
        (3, 31, "economy", datetime(2026, 10, 22, 18, 5), "booked"),    # Charlie: quoted 550,000, charged 1,000,000
        (5, 31, "galaxium", datetime(2026, 10, 27, 19, 42), "booked"),  # Eve: quoted 550,000, charged 5,000,000
        (7, 31, "business", datetime(2026, 10, 29, 12, 16), "booked"),  # Grace: quoted 2,750,000, charged 2,500,000
        (8, 29, "business", datetime(2026, 9, 18, 10, 31), "booked"),   # Heidi: the refund question
    ]
    taken: Counter = Counter()
    raw = []
    for user_id, flight_id, seat_class, when, status in fixed:
        taken[(flight_id, seat_class)] += 1
        raw.append((when, user_id, flight_id, seat_class, status))

    start = datetime(2026, 8, 1)
    while len(raw) < 340:
        when = start + timedelta(days=int(rng.triangular(0, 92, 80)), hours=rng.randint(6, 22), minutes=rng.randint(0, 59))
        open_flights = [f for f in flights if departs[f["flight_id"]] > when + timedelta(days=2)]
        flight = rng.choices(open_flights, weights=[ROUTES[(f["flight_id"] - 1) % 10][5] for f in open_flights])[0]
        seat_class = rng.choices(list(SEATS), weights=[60, 30, 10])[0]
        user_id = rng.randint(1, 10) if rng.random() < 0.3 else rng.randint(11, len(users))
        cancelled = rng.random() < 0.09
        if not cancelled:
            if taken[(flight["flight_id"], seat_class)] >= SEATS[seat_class]:
                continue
            taken[(flight["flight_id"], seat_class)] += 1
        status = "cancelled" if cancelled else ("completed" if departs[flight["flight_id"]] < TODAY else "booked")
        raw.append((when, user_id, flight["flight_id"], seat_class, status))

    raw.sort()
    bookings = [{
        "booking_id": i + 1, "user_id": u, "flight_id": fid, "seat_class": sc, "status": st,
        "booking_time": when.strftime("%Y-%m-%dT%H:%M:%SZ"), "price_paid": charged_price(by_id[fid], sc),
    } for i, (when, u, fid, sc, st) in enumerate(raw)]

    # Quotes exist only since the hold service went live. A converted quote precedes its
    # booking by a few minutes; the rest expired with their 15 minute hold.
    drafts = []
    for b in bookings:
        when = datetime.strptime(b["booking_time"], "%Y-%m-%dT%H:%M:%SZ")
        if when >= HOLD_SERVICE_LIVE:
            drafts.append((when - timedelta(minutes=rng.randint(2, 14)), b["user_id"], b["flight_id"], b["seat_class"], b["booking_id"]))
    for _ in range(len(drafts) * 2 // 5):
        when = HOLD_SERVICE_LIVE + timedelta(days=rng.randint(0, 30), hours=rng.randint(6, 22), minutes=rng.randint(0, 59))
        flight = rng.choice([f for f in flights if departs[f["flight_id"]] > when + timedelta(days=2)])
        drafts.append((when, rng.randint(1, len(users)), flight["flight_id"], rng.choices(list(SEATS), weights=[60, 30, 10])[0], ""))
    drafts.sort(key=lambda d: d[0])
    quotes = [{
        "quote_id": 5001 + i, "user_id": u, "flight_id": fid, "seat_class": sc, "quoted_price": quoted_price(fid, sc),
        "created_at": when.strftime("%Y-%m-%dT%H:%M:%SZ"), "status": "converted" if bid else "expired", "booking_id": bid,
    } for i, (when, u, fid, sc, bid) in enumerate(drafts)]

    for name, rows in (("users", users), ("flights", flights), ("bookings", bookings), ("quotes", quotes)):
        write_csv(name, rows)
    return users, flights, bookings, quotes


# ---------------------------------------------------------------- company templates

def build_templates() -> None:
    folder = COMPANY / "drive" / "templates"
    deck = Presentation()
    deck.core_properties.title = "Galaxium Travels deck template"
    deck.core_properties.created = deck.core_properties.modified = FIXED_TIME
    slides = [
        (0, "Title of the deck", "Department · Author · Date"),
        (1, "Summary", "The answer in one sentence.\nThe number that supports it, with its source.\nThe decision you are asking for."),
        (1, "Finding", "One finding per slide.\nState the figure and where it comes from.\nSay what it means for Galaxium."),
        (1, "Recommendation", "What we should do.\nWhat it costs and who owns it.\nWhat happens if we do nothing."),
        (1, "Sources and open points", "Every system or file a number came from.\nAnything that could not be verified."),
    ]
    for layout, title, body in slides:
        slide = deck.slides.add_slide(deck.slide_layouts[layout])
        slide.shapes.title.text = title
        slide.placeholders[1].text = body
        for paragraph in slide.placeholders[1].text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(20)
    deck.save(folder / "galaxium-deck-template.pptx")

    doc = Document()
    doc.core_properties.created = doc.core_properties.modified = FIXED_TIME
    doc.add_heading("Weekly bookings and cancellations report", 0)
    doc.add_paragraph("Week: <Monday> to <Sunday>, by booking date. Prepared by: <name>. Source: booking system.")
    doc.add_heading("Headline", 1)
    table = doc.add_table(rows=4, cols=3, style="Light Grid Accent 1")
    for i, row in enumerate([("Measure", "This week", "Previous week"), ("Bookings", "", ""), ("Cancellations", "", ""),
                             ("Revenue excluding cancelled bookings (credits)", "", "")]):
        for j, text in enumerate(row):
            table.cell(i, j).text = text
    doc.add_heading("Split by seat class", 1)
    doc.add_paragraph("<bookings and revenue for economy, business and galaxium>")
    doc.add_heading("Three strongest routes", 1)
    doc.add_paragraph("<route, bookings, revenue>")
    doc.add_heading("Anything unusual", 1)
    doc.add_paragraph("<one line, or 'nothing unusual'>")
    doc.save(folder / "weekly-bookings-report-template.docx")


# ---------------------------------------------------------------- pack: quote mismatch

def build_quote_mismatch(users: list[dict], flights: list[dict], bookings: list[dict], quotes: list[dict]) -> None:
    pack = PACKS / "quote-mismatch"
    by_flight = {f["flight_id"]: f for f in flights}
    name = {u["user_id"]: u["name"] for u in users}
    route = lambda fid: f"{by_flight[fid]['origin']} to {by_flight[fid]['destination']}"  # noqa: E731

    october = [b for b in bookings if b["booking_time"].startswith("2026-10")]
    export = [{**b, "customer": name[b["user_id"]], "route": route(b["flight_id"])} for b in october]
    # The planted error: one booking appears twice in the finance export.
    dup = next(b for b in export if b["booking_time"] >= "2026-10-15" and b["seat_class"] == "business" and b["status"] == "booked")
    with_dup = []
    for row in export:
        with_dup.append(row)
        if row is dup:
            with_dup.append(dict(row))
    write_xlsx(pack / "drive" / "finance" / "bookings-export-2026-10.xlsx", {"Bookings October 2026": with_dup})
    write_xlsx(pack / "drive" / "finance" / "quotes-export-2026-10.xlsx", {"Quotes October 2026": quotes})

    by_booking = {b["booking_id"]: b for b in bookings}
    pairs = [(q, by_booking[q["booking_id"]]) for q in quotes if q["status"] == "converted"]
    diffs = [(b["price_paid"] - q["quoted_price"], q, b) for q, b in pairs]
    more = [d for d in diffs if d[0] > 0]
    less = [d for d in diffs if d[0] < 0]
    equal = [d for d in diffs if d[0] == 0]
    per_class: dict[str, list] = defaultdict(list)
    for d in diffs:
        per_class[d[1]["seat_class"]].append(d)
    cancelled = [d for d in diffs if d[2]["status"] == "cancelled"]
    assert all(q["quoted_price"] == quoted_price(q["flight_id"], "economy") for q in quotes if q["seat_class"] == "galaxium")

    def case(user_id: int) -> str:
        d, q, b = next(x for x in diffs if x[1]["user_id"] == user_id and x[1]["flight_id"] == 31)
        return (f"| {name[user_id]} | {b['booking_id']} | {q['quote_id']} | {b['seat_class']} | "
                f"{money(q['quoted_price'])} | {money(b['price_paid'])} | {d:+,} |")

    lines = [
        "# Ground truth: the quote mismatch", "",
        "Generated by `scripts/build.py` from the snapshot. Do not edit by hand.", "",
        "## The short answer", "",
        f"- Between 1 and 31 October 2026 the hold service issued **{len(quotes)} quotes**. "
        f"**{len(pairs)}** became bookings and {len(quotes) - len(pairs)} expired.",
        f"- Only **{len(equal)} of {len(pairs)}** booked quotes match the amount charged.",
        f"- Customers were charged **more** than quoted on **{len(more)}** bookings (total {money(sum(d[0] for d in more))} credits) "
        f"and **less** than quoted on **{len(less)}** bookings (total {money(-sum(d[0] for d in less))} credits).",
        f"- Total quoted {money(sum(q['quoted_price'] for q, _ in pairs))}, total charged {money(sum(b['price_paid'] for _, b in pairs))}, "
        f"net difference **{money(sum(d[0] for d in diffs))} credits** charged above quote.",
        f"- {len(cancelled)} of the {len(pairs)} bookings were cancelled later. The figures above include them.", "",
        "## By seat class", "",
        "| Seat class | Booked quotes | Quoted | Charged | Charged minus quoted | Matching |", "| --- | ---: | ---: | ---: | ---: | ---: |",
    ]
    for sc in ("economy", "business", "galaxium"):
        rows = per_class[sc]
        lines.append(f"| {sc} | {len(rows)} | {money(sum(q['quoted_price'] for _, q, _ in rows))} | "
                     f"{money(sum(b['price_paid'] for _, _, b in rows))} | {sum(d for d, _, _ in rows):+,} | {sum(1 for d, _, _ in rows if d == 0)} |")
    lines += [
        "", "## The cause", "",
        "The charged amounts follow the pricing policy (base price of the flight times 1, 2.5 or 5). The quotes do not:", "",
        "- A quote ignores the route. It is 500,000 for economy and 2,500,000 for business, times 1.0, 1.1 or 1.2 depending on the flight id.",
        f"- The hold service does not know the galaxium class. All {sum(q['seat_class'] == 'galaxium' for q in quotes)} galaxium quotes equal the economy quote for the same flight.",
        "- So the charged amount is the correct one under the policy, and the quote is wrong. All three seat classes are affected.", "",
        "## The three cases named in mail and tickets", "",
        "All three are on flight 31, Earth to Mars, departing 3 December 2026.", "",
        "| Customer | Booking | Quote | Seat class | Quoted | Charged | Difference |", "| --- | --- | --- | --- | ---: | ---: | ---: |",
        case(3), case(5), case(7), "",
        "## Largest differences", "",
        "| Booking | Customer | Route | Seat class | Quoted | Charged | Difference |", "| --- | --- | --- | --- | ---: | ---: | ---: |",
    ]
    for d, q, b in sorted(diffs, key=lambda x: -abs(x[0]))[:5]:
        lines.append(f"| {b['booking_id']} | {name[b['user_id']]} | {route(b['flight_id'])} | {b['seat_class']} | "
                     f"{money(q['quoted_price'])} | {money(b['price_paid'])} | {d:+,} |")
    lines += [
        "", "## The planted error", "",
        f"`finance/bookings-export-2026-10.xlsx` lists **booking {dup['booking_id']} twice** ({dup['customer']}, {dup['route']}, "
        f"{dup['seat_class']}, {money(dup['price_paid'])}).", "",
        f"- Rows in the export: {len(with_dup)}. Bookings in October: **{len(october)}**.",
        f"- Sum of `price_paid` in the export as it is: {money(sum(r['price_paid'] for r in with_dup))}.",
        f"- Correct October total: **{money(sum(b['price_paid'] for b in october))}** "
        f"({money(sum(b['price_paid'] for b in october if b['status'] != 'cancelled'))} excluding cancelled bookings).", "",
        "A good answer notices the duplicate, or works from the booking system and reports that the export does not agree with it.", "",
        "## A claim to test", "",
        "In #eng-booking, Dev Malhotra says only the galaxium class is affected. That is false: economy and business quotes are wrong too, "
        "and business customers on cheaper routes were charged less than quoted.", "",
        "## What a good result contains", "",
        "1. The counts and the net difference above, with the source of each figure.",
        "2. The cause in plain words, and that the quote is wrong, not the charge.",
        "3. A memo to Ingrid Solberg in the drive, a ticket in the engineering queue for Marcus Lindqvist that references GT-1011, "
        "and a short post in #finance. All three show up on the Activity tab.",
        "4. The decision left to a person: whether to honour the quotes for customers who were charged more.",
    ]
    (pack / "ground-truth.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("quote-mismatch:", {name[u]: case(u) for u in (3, 5, 7)}, "dup", dup["booking_id"])


# ---------------------------------------------------------------- pack: win-loss

def build_win_loss() -> None:
    pack = PACKS / "win-loss"
    deals = yaml.safe_load((pack / "crm" / "deals.yaml").read_text(encoding="utf-8"))
    accounts = {a["id"]: a for a in yaml.safe_load((COMPANY / "crm" / "accounts.yaml").read_text(encoding="utf-8"))}
    flipped = "D-2014"  # the planted error: still "Won" in the manual tracker
    tracker = [{
        "Deal": d["id"], "Name": d["name"], "Account": accounts[d["account_id"]]["name"], "Segment": accounts[d["account_id"]]["segment"],
        "Owner": d["owner"], "Amount": d["amount"], "Closed": d["close_date"],
        "Outcome": "Won" if d["stage"] == "won" or d["id"] == flipped else "Lost",
        "Loss reason": "" if d["id"] == flipped else d.get("loss_reason", ""),
        "Competitor": "" if d["id"] == flipped else d.get("competitor", ""),
    } for d in deals]
    write_xlsx(pack / "drive" / "sales" / "pipeline-tracker-q3-2026.xlsx", {"Closed deals Q3 2026": tracker})

    won = [d for d in deals if d["stage"] == "won"]
    lost = [d for d in deals if d["stage"] == "lost"]
    total = lambda ds: sum(d["amount"] for d in ds)  # noqa: E731
    pct = lambda a, b: f"{100 * a / b:.1f}%"  # noqa: E731
    bad = next(d for d in deals if d["id"] == flipped)
    lines = [
        "# Ground truth: third quarter win-loss", "",
        "Generated by `scripts/build.py` from the CRM deals of this pack. Do not edit by hand.", "",
        "## The short answer", "",
        f"- **{len(deals)} deals** closed between 1 July and 30 September 2026: **{len(won)} won, {len(lost)} lost**.",
        f"- Win rate by count: **{pct(len(won), len(deals))}**. Win rate by value: **{pct(total(won), total(deals))}**.",
        f"- Won value {money(total(won))} credits. Lost value {money(total(lost))} credits.",
        f"- Average won deal {money(total(won) // len(won))}. Average lost deal {money(total(lost) // len(lost))}. We lose the larger deals.", "",
        "## Why we lose", "",
        "| Loss reason | Deals | Value | Share of lost value |", "| --- | ---: | ---: | ---: |",
    ]
    reasons: dict[str, list] = defaultdict(list)
    for d in lost:
        reasons[d["loss_reason"]].append(d)
    for reason, ds in sorted(reasons.items(), key=lambda kv: -total(kv[1])):
        lines.append(f"| {reason} | {len(ds)} | {money(total(ds))} | {pct(total(ds), total(lost))} |")
    capacity = reasons["Schedule frequency"] + reasons["Outer-system coverage"]
    lines += [
        "", f"Schedule frequency and outer-system coverage together are **{len(capacity)} of {len(lost)} losses** and "
        f"**{pct(total(capacity), total(lost))} of lost value**. Price accounts for {len(reasons['Price'])} losses "
        f"({pct(total(reasons['Price']), total(lost))} of lost value).", "",
        "## By segment", "", "| Segment | Won | Lost | Win rate | Won value | Lost value |", "| --- | ---: | ---: | ---: | ---: | ---: |",
    ]
    for seg in sorted({a["segment"] for a in accounts.values()}):
        w = [d for d in won if accounts[d["account_id"]]["segment"] == seg]
        lo = [d for d in lost if accounts[d["account_id"]]["segment"] == seg]
        lines.append(f"| {seg} | {len(w)} | {len(lo)} | {pct(len(w), len(w) + len(lo))} | {money(total(w))} | {money(total(lo))} |")
    lines += ["", "## By competitor", "", "| Competitor | Deals lost to them | Value |", "| --- | ---: | ---: |"]
    for comp, n in Counter(d["competitor"] for d in lost).most_common():
        lines.append(f"| {comp} | {n} | {money(total([d for d in lost if d['competitor'] == comp]))} |")
    lines += ["", "## By owner", "", "| Owner | Won | Lost | Win rate |", "| --- | ---: | ---: | ---: |"]
    for owner in sorted({d["owner"] for d in deals}):
        w, lo = sum(d["owner"] == owner for d in won), sum(d["owner"] == owner for d in lost)
        lines.append(f"| {owner} | {w} | {lo} | {pct(w, w + lo)} |")
    lines += [
        "", "## The planted error", "",
        f"`sales/pipeline-tracker-q3-2026.xlsx` still lists **{flipped} ({bad['name']}, {money(bad['amount'])})** as Won. "
        "The CRM says lost, and the call note of 11 September explains why. The CRM is the system of record.", "",
        f"- Tracker as it is: {len(won) + 1} won, {len(lost) - 1} lost, win rate {pct(len(won) + 1, len(deals))}, "
        f"won value {money(total(won) + bad['amount'])}.",
        f"- Correct: {len(won)} won, {len(lost)} lost, win rate {pct(len(won), len(deals))}, won value {money(total(won))}.", "",
        "## A claim to test", "",
        "In #sales, Dev Malhotra says the team mostly loses on price. The CRM shows 2 of 10 losses on price. The notes of two lost deals "
        "say that price was not the issue (D-2002, D-2007), and two deals were won at a higher price than the competitor's (D-2001, D-2008).", "",
        "## What a good result contains", "",
        "1. The win rate by count and by value, from the CRM and not from the tracker, with the discrepancy on D-2014 pointed out.",
        "2. The loss reasons ranked by value, with the conclusion that capacity (frequency and outer-system coverage) is the problem, not price.",
        "3. A deck built on `templates/galaxium-deck-template.pptx` with sources on the last slide.",
        "4. A post in #sales with the headline numbers. Optional: the tracker error noted on the D-2014 deal in the CRM.",
    ]
    (pack / "ground-truth.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


# ---------------------------------------------------------------- pack: vendor contract

CLAUSES = [
    ("1. Services", "Tycho Ground Services Ltd. (\"Vendor\") provides ground handling for Galaxium Travels (\"Customer\") departures and arrivals at Tycho Spaceport: docking, fuelling, baggage and passenger transfer."),
    ("2. Term and renewal", "This Agreement runs for an initial term of 36 months from 1 January 2027. It renews automatically for successive periods of 24 months unless either party gives written notice at least 180 days before the end of the current term."),
    ("3. Fees", "The Customer pays a monthly fee of 420,000 credits for up to 16 turnarounds per month. Each further turnaround is charged at 30,000 credits."),
    ("4. Price adjustment", "The Vendor may increase all fees once per contract year by up to 5%."),
    ("5. Payment", "Invoices are issued monthly in advance and are payable within 15 days of the invoice date."),
    ("6. Service levels", "The Vendor will use reasonable efforts to complete each turnaround within 120 minutes of docking. Failure to meet this target does not entitle the Customer to any credit, refund or other remedy."),
    ("7. Confidentiality", "Each party keeps the other party's confidential information secret for the term and for three years after it ends."),
    ("8. Limitation of liability", "The total liability of each party under this Agreement in any contract year is limited to the fees paid in the three months before the event giving rise to the claim."),
    ("9. Termination", "Either party may terminate this Agreement for material breach that is not remedied within 30 days of written notice. The Vendor may terminate this Agreement for convenience with 60 days' written notice."),
    ("10. Incident notification", "The Vendor notifies the Customer of any safety or data incident affecting the Customer's passengers or operations within 48 hours of becoming aware of it."),
    ("11. Governing law", "This Agreement is governed by the law of the Lunar Free Port."),
    ("12. Entire agreement", "This Agreement replaces the Ground Handling Agreement of 1 January 2025 between the parties."),
]


def build_vendor_contract() -> None:
    pack = PACKS / "vendor-contract"
    doc = Document()
    doc.core_properties.created = doc.core_properties.modified = FIXED_TIME
    doc.add_heading("Ground Handling Agreement", 0)
    doc.add_paragraph("between Tycho Ground Services Ltd. and Galaxium Travels. Draft of 30 October 2026, prepared by the Vendor.")
    for title, text in CLAUSES:
        doc.add_heading(title, 2)
        doc.add_paragraph(text)
    target = pack / "drive" / "legal" / "tycho-ground-handling-agreement-draft.docx"
    target.parent.mkdir(parents=True, exist_ok=True)
    doc.save(target)

    (pack / "ground-truth.md").write_text("""# Ground truth: the Tycho ground handling agreement

Written together with the contract in `scripts/build.py`. Do not edit by hand.

## Deviations from the vendor contract playbook

Seven clauses deviate. All seven are in the "escalate" column of the playbook.

| Clause | The draft says | Playbook standard | Fallback | Verdict |
| --- | --- | --- | --- | --- |
| 2. Term and renewal | 36 months, automatic renewal for 24 months, 180 days' notice | Up to 24 months, no automatic renewal | 12 month renewal, 90 days' notice | Escalate: term, renewal period and notice period all exceed the fallback |
| 4. Price adjustment | Up to 5% a year | Fixed for the term | Capped at 3% | Escalate |
| 5. Payment | In advance, 15 days | Net 45 | Net 30 | Escalate |
| 6. Service levels | 120 minutes, reasonable efforts, no credits | 90 minutes, credits of at least 5% | Credits of at least 3% | Escalate: slower target and no service credits |
| 8. Limitation of liability | 3 months of fees | 12 months of fees | 9 months | Escalate |
| 9. Termination | Only the Vendor may terminate for convenience | Galaxium may terminate with 60 days' notice | 90 days | Escalate: Galaxium has no such right |
| 3. Fees | 420,000 a month, up from 380,000 | Not a playbook topic | | Flag: a 10.5% increase, check against the proposal |

## Clauses that are acceptable

- **10. Incident notification, 48 hours.** Not the standard (24 hours) but within the fallback. No escalation.
- **11. Governing law, Lunar Free Port.** Within the fallback. No escalation.
- 1, 7 and 12 raise no playbook issue.

A good review separates these from the escalations. Reporting clauses 10 and 11 as violations is a mistake.

## The planted error

The vendor's cover mail of 30 October says the liability cap is "unchanged at twelve months of fees". Clause 8 of the attached draft says **three months**. The mail is wrong about the document it introduces.

## The link to operations

In October, 3 of 14 turnarounds at Tycho were later than 90 minutes (2, 9 and 21 October), according to the vendor's own mail to Priya Raman. Under the playbook standard each would earn a credit of at least 5% of the monthly fee. Under clause 6 of the draft Galaxium would get nothing, and the target would relax to 120 minutes.

## The other proposal

Clavius Port Services offers 395,000 credits a month, a 24 month term without automatic renewal, 90 minute turnarounds with 5% credits, and a liability cap of 12 months. Clavius relies on the shared fuelling gantry until July 2027, while Tycho's second gantry is due in January. Clavius meets the playbook on every topic it mentions; payment terms and termination are not stated and must be asked.

## What a good result contains

1. A review memo in the drive that follows the playbook's format: clause, what the draft says, what the playbook says, proposed change.
2. The seven deviations, with 10 and 11 marked as acceptable fallbacks.
3. The contradiction between the cover mail and clause 8.
4. A ticket in the legal queue or a mail to Helena Vogt asking for the escalation decision.
""", encoding="utf-8")


if __name__ == "__main__":
    data = build_snapshot()
    build_templates()
    build_quote_mismatch(*data)
    build_win_loss()
    build_vendor_contract()
    print(f"snapshot: {len(data[0])} users, {len(data[1])} flights, {len(data[2])} bookings, {len(data[3])} quotes")
