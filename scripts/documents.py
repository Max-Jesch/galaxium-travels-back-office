"""The decks, workbooks and documents a company accumulates: board packs, briefings,
registers, plans, contracts, guides.

`routine.build()` calls `build()` after its own drive files. Everything lands in
`content/generated/drive/`. The documents about the ships, the ports and the cabin are
in `worldbook.py`, written from `content/company/fleet.yaml`, `ports.yaml` and
`cabin.yaml`; the same rules apply there. The rules for what may be said here:

- Bookings, quotes and flights are read from the snapshot, deals from the CRM file.
  Nothing about them is typed in.
- Figures that already appear in a hand-written note, mail or chat are repeated exactly
  (the support weekly numbers, the Tycho turnarounds, the Halcyon repair, the budgets).
- Nothing here answers a question that a scenario pack or an open mail asks: no
  third-quarter win-loss, no comparison of quotes with charges, no review of the Tycho
  draft, no Europa load, no Pluto options beyond the outline.
- Every document is dated on or before the story's today, 4 November 2099.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import yaml

import companybook
import worldbook
from office import AMBER, OK, WARN, Book, Brand, Deck, Doc, F

ROOT = Path(__file__).resolve().parent.parent
DRIVE = ROOT / "content" / "generated" / "drive"
COMPANY = ROOT / "content" / "company"
MONTH = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
DAY = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

TYCHO = Brand("TYCHO GROUND SERVICES LTD.", primary="1F3B4D", accent="3A8FA3", head="Arial", text="Arial",
              footer="Tycho Ground Services Ltd. · Tycho Spaceport, Luna · tychoground.example")
MARLOWE = Brand("MARLOWE & FINCH · AUDITORS", primary="1F3D2B", accent="9A7B3C", head="Georgia", text="Georgia",
                footer="Marlowe & Finch, Auditors · marlowefinch.example")
AGREEMENT = Brand("", primary="1B1A2E", accent="6B6A80", head="Georgia", text="Georgia", footer="")

# Repeated from support/weekly-numbers.md.
SUPPORT_WEEKS = [("28 Sep", 171, 36, 4.5, "Refund amount"), ("5 Oct", 184, 38, 4.5, "Hold expired"), ("12 Oct", 231, 52, 4.3, "Payment timeouts (13 Oct)"),
                 ("19 Oct", 176, 35, 4.5, "Hold expired"), ("26 Oct", 198, 41, 4.4, "Name does not match")]

# The safety reports of the third quarter and of October so far. The log and the pre-read are both built from this list.
SAFETY = [
    ("SR-2099-031", "2099-07-08", "Cape Meridian", "Ground", "Fuel hose still connected after the handler's sign-off. Found on the walk-round.", "Low", "Closed"),
    ("SR-2099-032", "2099-07-21", "Solstice", "Cabin", "Galley cart brake did not hold during boarding.", "Low", "Closed"),
    ("SR-2099-033", "2099-08-03", "Olympus Port", "Ground", "Walkway lighting out on the east side of pad 1.", "Low", "Closed"),
    ("SR-2099-034", "2099-08-14", "Aurora", "Crew", "Crew member reported fatigue before a Jupiter leg. Replaced from standby, departure on time.", "Medium", "Closed"),
    ("SR-2099-035", "2099-08-27", "Tycho", "Ground", "Baggage loader reversed without a second person watching.", "Medium", "Closed"),
    ("SR-2099-036", "2099-09-06", "Kestrel", "Cabin", "Overhead restraint latch opened during the departure burn, row 2. Bag stayed in place.", "Medium", "Open: bulletin 2099-07"),
    ("SR-2099-037", "2099-09-19", "Kestrel", "Cabin", "Overhead restraint latch opened during the departure burn, row 5. Bag stayed in place.", "Medium", "Open: bulletin 2099-07"),
    ("SR-2099-038", "2099-09-23", "Aphrodite Station", "Ground", "Pressure seal warning at berth 3 on docking. Cleared after the seal was reseated.", "Low", "Closed"),
    ("SR-2099-039", "2099-10-17", "Tycho", "Passenger", "Booked wheelchair assistance not provided on arrival. Cabin crew assisted. Third case in October (ticket GT-1019).", "Medium", "Closed 20 October"),
    ("SR-2099-040", "2099-10-24", "Halcyon", "Technical", "Wear on the number 2 coolant valve at the routine inspection. Inside limits. Assembly replaced 26 to 29 October.", "Medium", "Closed 29 October"),
]


def path(rel: str) -> Path:
    return DRIVE / rel


def long(d: date | str) -> str:
    d = date.fromisoformat(str(d)[:10])
    return f"{d.day} {MONTH[d.month]} {d.year}"


def short(d: date | str) -> str:
    d = date.fromisoformat(str(d)[:10])
    return f"{d.day} {MONTH[d.month][:3]}"


class World:
    """The snapshot and the CRM, with the few questions the documents ask of them."""

    def __init__(self, d) -> None:
        self.d, self.people = d, d.people
        self.flights, self.bookings = d.flights, d.bookings
        self.deals = yaml.safe_load((COMPANY / "crm" / "deals.yaml").read_text(encoding="utf-8"))
        self.accounts = {a["id"]: a for a in yaml.safe_load((COMPANY / "crm" / "accounts.yaml").read_text(encoding="utf-8"))}
        self.open_deals = sorted((x for x in self.deals if x["stage"] not in ("won", "lost")), key=lambda x: str(x["close_date"]))

    def booked(self, first: str, last: str) -> list[dict]:
        return [b for b in self.bookings if first <= b["booking_time"][:10] <= last]

    def sold(self, flight_id: int, as_of: str) -> int:
        return sum(b["flight_id"] == flight_id and b["status"] != "cancelled" and b["booking_time"][:10] <= as_of for b in self.bookings)

    def hours(self, f: dict) -> int:
        fmt = "%Y-%m-%dT%H:%M:%SZ"
        return int((datetime.strptime(f["arrival_time"], fmt) - datetime.strptime(f["departure_time"], fmt)).total_seconds() // 3600)

    def account(self, deal: dict) -> str:
        return self.accounts[deal["account_id"]]["name"]

    def at(self, location: str) -> list[dict]:
        return [p for p in self.people if p.get("location") == location]


# ====================================================================== decks

def decks(w: World) -> None:
    d = w.d

    # ---- Board pack, August. The layout Ingrid asks to be reused in November. Its source lines are vague on purpose:
    # the board asked where the numbers came from.
    aug = w.booked("2099-08-01", "2099-08-16")
    by_route, by_class = defaultdict(list), defaultdict(list)
    for b in aug:
        by_route[d.route(b["flight_id"])].append(b)
        by_class[b["seat_class"]].append(b)
    rows = lambda groups: [[k, len(v), sum(x["status"] == "cancelled" for x in v), sum(x["price_paid"] for x in v)] for k, v in sorted(groups.items())]  # noqa: E731
    total = lambda rs: [["Total", sum(r[1] for r in rs), sum(r[2] for r in rs), sum(r[3] for r in rs)]]  # noqa: E731
    q2 = [x for x in w.deals if "2099-04-01" <= str(x["close_date"]) <= "2099-06-30" and x["stage"] in ("won", "lost")]
    deck = Deck(path("leadership/board/board-pack-2099-08.pptx"), "Board pack, August 2099", "Galaxium Travels · Finance · 20 August 2099",
                owner="Ingrid Solberg", modified="2099-08-19")
    deck.bullets("Summary", ["**Scheduled service** on ten routes, one departure a month each.", "**Pluto**, added in the spring, fills steadily.",
                             "**Hold service** on track for 1 October: travellers see a price and have 15 minutes to pay.",
                             "**Fleet:** five ships in service. Perihelion is in refit until February 2100."],
                 aside=("Asked of the board", ["No decision this month.", "Next meeting: 19 November 2099."]))
    deck.table("Scheduled service: bookings 1 to 16 August, by route", ["Route", "Bookings", "Cancelled", "Charged (dollars)"], rows(by_route) + total(rows(by_route)),
               widths=[3, 1, 1, 2], total=True, source="Finance", note="Bookings by booking time. The full month follows with the August close.")
    deck.table("Scheduled service: bookings 1 to 16 August, by seat class", ["Seat class", "Bookings", "Cancelled", "Charged (dollars)"],
               rows(by_class) + total(rows(by_class)), widths=[3, 1, 1, 2], total=True, source="Finance",
               note="Each departure has ten seats: 6 economy, 3 business, 1 galaxium.")
    deck.table("Charter: contracts decided in the second quarter", ["Deal", "Account", "Owner", "Outcome", "Value (dollars)"],
               [[x["id"], w.account(x), x["owner"], x["stage"].capitalize(), x["amount"]] for x in sorted(q2, key=lambda x: str(x["close_date"]))],
               widths=[1, 3.2, 1.6, 1, 1.6], source="Sales", note="Charter flying is sold under contract and never appears in the booking system.")
    cash = [["Operating account", 1096200000], ["Payroll account", 41300000], ["Luna entity", 28300000], ["Maintenance reserve", 18500000]]
    deck.table("Cash at 31 July", ["Account", "Balance (dollars)"], cash + [["Total", sum(r[1] for r in cash)]], widths=[3, 2], total=True, source="Finance")
    ageing = [["Not yet due", 14820000], ["1 to 30 days overdue", 3260000], ["31 to 60 days overdue", 410000], ["More than 60 days overdue", 0]]
    deck.table("Payables ageing at 31 July", ["Age", "Open (dollars)"], ageing + [["Total", sum(r[1] for r in ageing)]], widths=[3, 2], total=True, source="Finance")
    deck.table("Fleet", ["Ship", "Class", "Certified to", "Status"],
               [["Kestrel", "Kestrel", "Mars and Venus", "In service"], ["Meridian", "Kestrel", "Mars and Venus", "In service"], ["Solstice", "Kestrel", "Mars and Venus", "In service"],
                ["Aurora", "Aurora", "Jupiter", "In service"], ["Halcyon", "Halcyon", "Pluto", "In service. The only outer-system hull."],
                ["Perihelion", "Halcyon", "Pluto after refit", "Refit until February 2100"]], widths=[1.3, 1.2, 1.6, 3.2], source="Operations")
    deck.bullets("Risks", ["**One outer-system hull.** Until Perihelion returns there is no slack beyond Jupiter. Owner: Olu Adeyemi.",
                           "**Two price calculations** once the hold service is live. They are to be merged. Owner: Farid Nazari.",
                           "**Manual exports.** Finance and Operations pull their numbers by hand. Owner: Farid Nazari."])
    deck.statement("Decisions asked of the board", "None this month.", "The November meeting takes the 2100 budget envelope and the Luna ground handling agreement, which ends on 31 December 2099.")
    deck.save()

    # ---- All-hands, 1 October
    weeks, monday = [], date(2099, 8, 3)
    while monday <= date(2099, 9, 21):
        weeks.append((f"{monday.day} {MONTH[monday.month][:3]}", len(w.booked(str(monday), str(monday + timedelta(days=6))))))
        monday += timedelta(days=7)
    n_aug, n_sep = len(w.booked("2099-08-01", "2099-08-31")), len(w.booked("2099-09-01", "2099-09-30"))
    deck = Deck(path("leadership/all-hands/all-hands-2099-10-01.pptx"), "All-hands", "Thursday 1 October 2099 · Atrium, Earth HQ, and video",
                owner="Renata Castellanos", modified="2099-10-01")
    deck.bullets("Today", ["The third quarter in one chart", "The hold service, live since this morning", "The fleet until February", "What the fourth quarter is about", "Your questions"])
    deck.chart("Bookings on the scheduled service grew through the quarter", "column", [k for k, _ in weeks], {"Bookings in the week starting": [v for _, v in weeks]},
               source="Booking system, bookings by booking date, pulled 1 October 2099",
               aside=("What it says", [f"August: {n_aug} bookings.", f"September: {n_sep} bookings, the strongest month so far.", "Charter flying is not in this chart. It is sold under contract and lives in the CRM."]))
    deck.timeline("The hold service is live", [("Search", "The traveller picks a departure and a seat class."), ("Quote", "We show a price."),
                                               ("Hold", "The seat is held for 15 minutes."), ("Pay", "Paid inside the hold: the quote becomes a booking."),
                                               ("Or not", "The hold runs out, the seat goes back on sale, nothing is charged.")],
                  items=["Live for all customers since this morning.", "A quote is not a booking. It becomes one when the traveller pays inside the hold.",
                         "Support has a macro for \"my hold expired\". Thank you, Greta."])
    deck.bullets("The fleet until February", ["**Kestrel, Meridian and Solstice** fly the inner routes.", "**Aurora** flies Mars and Jupiter.",
                                              "**Halcyon** flies Europa and Pluto, alone.", "**Perihelion** returns from refit in February and becomes the second outer-system hull."],
                 aside=("Why it matters", ["With one hull beyond Jupiter, every finding at an inspection is a schedule question."]))
    deck.two("The fourth quarter", ("For travellers and customers", ["Travellers know the price before they pay.", ("Fewer than 1% of bookings with a price complaint.", 1),
                                                                     "Keep corporate customers.", ("Renew Outer Belt Logistics and Nimbus Freight.", 1)]),
             ("For ourselves", ["Decide the 2100 fleet.", ("Options paper for a second Pluto departure by 9 November.", 1),
                                "Stop running the company from pasted spreadsheets.", ("Monthly finance exports written by a job from the November close.", 1)]),
             source="leadership/okrs-2099-q4.md")
    deck.two("You asked", ("Will Luna get a canteen?", ["No. The lunch allowance goes up instead. (Dorothea)"]),
             ("When do we stop exporting by hand?", ["Felix has started. Finance first. (Farid)"]))
    deck.statement("Next all-hands", "Thursday 5 November, 15:00.", "Questions to Agnes beforehand, or on the day.")
    deck.save()

    # ---- Winter on Mars briefing
    mars = [f for f in w.flights.values() if f["destination"] == "Mars" and "2099-12-01" <= f["departure_time"][:10] <= "2100-02-28"]
    deck = Deck(path("marketing/campaigns/winter-on-mars/briefing-deck.pptx"), "Winter on Mars", "Campaign briefing · Marketing · October 2099",
                owner="Elise Fontaine", modified="2099-10-07")
    deck.statement("The idea", "Mars has a winter. The dune lodges at Red Sands are at their best in it.", "We sell the trip. Red Sands sells the stay.")
    deck.bullets("Who it is for", ["Leisure travellers from Earth and Luna who have flown with us once.", "Small groups.",
                                   "Not a first-flight campaign: these people know what a departure with us is like."],
                 aside=("What success looks like", ["Earth to Mars and Moon to Mars departures in January and February sold out four weeks ahead."]))
    deck.timeline("Dates", [("10 November", "Launch, 09:00."), ("10 December", "First departures at the group rate."), ("28 February", "Campaign and group rate end.")])
    deck.table("The offer", ["Where", "Price", "Conditions"],
               [["Public booking site", "No discount", "Prices follow the pricing policy."],
                ["Groups of eight or more, through Sales", "10% off", "Departures 10 December to 28 February. At most 40 seats a month at that rate."]],
               widths=[2.2, 1.2, 4], source="marketing/campaigns/winter-on-mars/campaign-brief.md", size=14)
    deck.table("Scheduled departures to Mars in the campaign", ["Flight", "Route", "Departure (UTC)", "Seats"],
               [[d.gx(f["flight_id"]), d.route(f["flight_id"]), f["departure_time"].replace("T", " ")[:16], 10] for f in mars],
               widths=[1, 2, 2, 1], source="Booking system, schedule", note="Ten seats per departure: 6 economy, 3 business, 1 galaxium. A group of eight takes most of a ship.")
    deck.bullets("Channels", ["Customer newsletter", "Paid media on Earth and Luna", "The Kepler Media films", "Our stand at the Interplanetary Travel Expo, 24 to 26 November"])
    deck.two("What we say", ("Say", ["\"Mars has a winter.\"", "Departure, traveller, seat hold.", "Before launch, to a customer who asks: \"Soon.\""]),
             ("Do not say", ["\"Guaranteed aurora.\" Legal has ruled it out, and so has astronomy.", "Blast off, adventure of a lifetime.", "Dates or prices before 10 November."]),
             source="marketing/brand/brand-guidelines.md")
    deck.table("Who does what", ["Item", "Owner"], [["Landing page copy, press note, newsletter", "Noor Rahimi"], ["Visuals", "Viktor Hale"], ["Group rate sheet with Red Sands", "Pilar Navarro"],
                                                    ["Paid media", "Elise Fontaine"], ["Support questions and answers", "Greta Olsen"], ["Kepler Media films", "Dev Malhotra, Kwame Asante"]],
               widths=[3, 2], size=14)
    deck.save()

    # ---- Corporate overview, for customers
    routes = sorted({(f["origin"], f["destination"], w.hours(f)) for f in w.flights.values()}, key=lambda r: (r[2], r[0]))
    places = sorted({p for r in routes for p in r[:2]})
    deck = Deck(path("sales/corporate-overview-deck.pptx"), "Galaxium Travels for corporate travel", "Crew rotations, research seasons and productions · 2099",
                owner="Amara Okafor", modified="2099-10-05")
    deck.statement("Who we are", "Interplanetary travel, booked like a train ticket.", "Scheduled service on ten routes, and charter rotations built around your calendar.")
    deck.figures("At a glance", [("6", "ships"), ("10", "scheduled routes"), (str(len(places)), "destinations, Earth to Pluto"), ("3", "stations with our own staff")],
                 items=["Earth HQ at Cape Meridian, an office on Luna at Tycho and one on Mars at Olympus Port.", "A named account executive and a customer success manager for every contract."])
    deck.table("Where we fly", ["From", "To", "Flight time", "Scheduled service"], [[o, dst, f"{h} hours", "One departure a month"] for o, dst, h in routes],
               widths=[1.5, 1.5, 1.2, 2], note="Charter rotations fly on your dates, on any of these routes.")
    deck.chart("A small ship, a quiet cabin", "doughnut", ["Economy", "Business", "Galaxium"], {"Seats": [6, 3, 1]}, labels=False,
               aside=("Ten seats per departure", ["6 economy, 3 business, 1 galaxium.", "A whole ship is a crew, a field team or a production unit.", "Whole-ship charters are priced as ten business seats less 20%."]))
    deck.two("Two ways to fly with us", ("Scheduled service", ["One departure per route per month.", "Booked on the public site, one seat at a time.", "A seat hold of 15 minutes: see the price, then decide."]),
             ("Charter rotations", ["Whole-ship or block bookings under a contract.", "Fixed dates that match your shift calendar.", "Invoices the way your finance team wants them: by site, by month, by traveller."]))
    deck.table("Who we fly for", ["Segment", "Typical need", "What we build"],
               [["Mining and Resources", "Crew rotations on fixed cycles", "Departure days that match the shift calendar"],
                ["Logistics", "Crew changes", "Block charters on crew change days"],
                ["Research and Education", "Seasons and field trips on fixed budgets", "One contract for the season, travel credit instead of lost fares"],
                ["Hospitality and Media", "Guests and productions", "Dates around the event, galaxium class for the top tier"]], widths=[2, 2.6, 3.4], size=13)
    deck.timeline("How a rotation is set up", [("1", "You tell us how often your crews rotate."), ("2", "Operations confirms the dates. We do not promise a departure before that."),
                                              ("3", "Contract and rate, from the rate card."), ("4", "Names at booking, as on the travel document."), ("5", "We fly. Your customer success manager reviews every quarter.")])
    deck.statement("Next step", "Tell us how often your crews rotate.", "We build the schedule around it.  sales@galaxium.example")
    deck.save()

    # ---- Booking roadmap
    deck = Deck(path("engineering/roadmap/booking-roadmap-2100-q1.pptx"), "Booking product roadmap", "First quarter 2100 · Product · 23 October 2099",
                owner="Maya Goldberg", modified="2099-10-23")
    deck.statement("Where we are going", "One price calculation, one refund answer, and holds that work for a group.", "Three weeks after the hold service went live.")
    deck.table("The list", ["Item", "Issue", "Why", "When"],
               [["Refund preview", 43, "Half of refund contacts are \"how much do I get back\"", "4.12, 10 November"],
                ["Quote price from the booking API", 51, "One price calculation instead of two", "4.13, December"],
                ["Group holds", 47, "Most requested by corporate bookers", "Q1"],
                ["Through-tickets with a connection", 58, "Jupiter to Europa to Earth is sold as two bookings", "Q1"],
                ["Status page", 60, "From the October postmortem", "November"],
                ["Monthly exports as a job", 62, "Finance and Operations pull them by hand", "November"]],
               widths=[2.6, 0.7, 4, 1.6], source="engineering/roadmap/roadmap-2100-q1.md", size=13)
    deck.timeline("In order", [("10 November", "Release 4.12: refund preview."), ("November", "Status page. Monthly exports as a job, finance first."),
                               ("December", "Release 4.13: the quote price comes from the booking API."), ("Q1 2100", "Group holds. Through-tickets with a connection.")])
    deck.chart("What Support hears", "column", [k for k, *_ in SUPPORT_WEEKS[:3]], {"Contacts in the week starting": [v for _, v, *_ in SUPPORT_WEEKS[:3]]},
               source="support/weekly-numbers.md", aside=("Top reason each week", [f"{k}: {reason}" for k, _, _, _, reason in SUPPORT_WEEKS[:3]]))
    deck.two("What corporate bookers ask for", ("Asked", ["A hold longer than 15 minutes for a group (Tharsis Engineering, 19 October).", "Names 72 hours before departure instead of at booking (same request)."]),
             ("Our answer today", ["Group holds are planned for early 2100. Until then Sales reserves group blocks by hand.", "Names stay required at booking: the traveller must match the account. Changes are cancel and rebook."]),
             source="Mail thread \"Request for quote, 40 rotations Mars to Jupiter in 2100\"")
    deck.bullets("Not on this roadmap", ["Discounts on the public site. Its prices follow the pricing policy.", "Charter flying. It is not in the booking system at all.",
                                         "Changing a booking in place. Cancel and rebook stays."])
    deck.save()

    # ---- Postmortem review
    deck = Deck(path("engineering/postmortems/2099-10-13-payment-timeouts-review.pptx"), "Payment timeouts, 13 October", "Postmortem review · Friday 16 October 2099 · Room Vesta",
                owner="Emeka Obi", modified="2099-10-16")
    deck.figures("What happened, in numbers", [("45 min", "15:50 to 16:35 UTC"), ("37", "payments timed out"), ("0", "customers charged twice"), ("23", "customers got two confirmation mails"), ("47", "contacts handled by Support")],
                 items=["For 45 minutes the booking API could not confirm payments: every connection to the payment provider was waiting on a slow answer."],
                 source="engineering/postmortems/2099-10-13-payment-timeouts.md")
    deck.timeline("Timeline (UTC)", [("15:45", "The provider's authorisations slow down: a failed database failover on their side."), ("15:50", "Our 20 connections to the provider are all waiting."),
                                     ("15:52", "Alert on response time. Incident opened."), ("16:22", "Pool raised to 60, timeout 10 seconds. Deployed as 4.11.1."),
                                     ("16:35", "Response times normal."), ("17:05", "Incident closed.")], source="Incident channel and the provider's report of 15 October")
    deck.bullets("Three causes", ["**The provider was slow.** Authorisations took up to 25 seconds.", "**Our pool was small and our timeout long**, so a slow provider stopped every booking.",
                                  "**The confirmation mail job retried after a timeout** without checking whether the first mail had gone."],
                 aside=("No blame", ["The pool size had been the same since 2097, and nobody had a reason to look at it."]))
    deck.table("Actions", ["Action", "Owner", "Due", "Status"], [["Alert on connection pool use", "Jonas Keller", "14 October", "Done"], ["Fix the mail retry", "Jonas Keller", "14 October", "Done (4.11.2)"],
                                                                 ["Circuit breaker for the payment provider", "Jonas Keller", "Release 4.12", "In progress"],
                                                                 ["Status page that Support can point customers to", "Emeka Obi", "November", "Open"]], widths=[4, 1.6, 1.4, 1.5], size=14)
    deck.bullets("What went well", ["Support told customers within minutes that retrying was safe, and it was.", "No payment was lost or duplicated.", "The fix was deployed 30 minutes after the alert."])
    deck.save()

    # ---- Welcome deck
    by_place, by_dept = Counter(p.get("location") for p in w.people), Counter(p["department"] for p in w.people)
    deck = Deck(path("people/onboarding/welcome-to-galaxium.pptx"), "Welcome to Galaxium Travels", "Your first week · People team", owner="Samir Chaudhry", modified="2099-10-02")
    deck.statement("What we do", "We carry people across the solar system.", "They trust us with their safety, their money and their plans.")
    deck.figures("Where we work", [(str(by_place["Earth HQ"]), "Earth HQ, Cape Meridian"), (str(by_place["Luna Station office"]), "Luna Station office, Tycho"), (str(by_place["Mars office"]), "Mars office, Olympus Port")],
                 items=["Core hours are 10:00 to 15:00 in your own station's time.", "Three stations, many time zones: do not schedule meetings at somebody else's night."],
                 source="people/org/headcount-2099-10.xlsx")
    dept = sorted(by_dept.items(), key=lambda kv: -kv[1])
    deck.chart("Who we are", "bar", [k for k, _ in dept], {"People": [v for _, v in dept]}, source="people/org/headcount-2099-10.xlsx",
               aside=("Office staff", [f"{len(w.people)} people in ten teams.", "Flight and cabin crew are rostered by Operations and are not in this count."]))
    deck.table("Three kinds of flying", ["Kind", "What it is", "Where the numbers are"],
               [["Scheduled service", "One departure per route per month on ten routes. Ten seats.", "The booking system"],
                ["Charter rotations", "Whole-ship or block bookings for corporate accounts.", "The CRM and Finance"], ["Shuttles", "Crew and positioning hops between Earth and Luna. Not sold.", "Operations control log"]],
               widths=[1.6, 4, 2], source="ops/how-we-fly.md", size=14,
               note="When somebody says \"bookings\", they mean the scheduled service. \"A deal\" is charter flying and never appears in the booking system.")
    deck.timeline("Your first week", [("Day 1", "Laptop and badge from the IT desk, third floor. Lunch with your manager. Enrol a second factor."), ("Day 2", "The handbook, the code of conduct and \"How we fly\"."),
                                      ("Day 3", "Security and privacy training, 25 minutes."), ("Day 4", "An hour with Support, listening to calls. Everyone does this."), ("Day 5", "Tell your manager what surprised you.")])
    deck.two("How we work", ("Write things down", ["Decisions live in the drive, not in a call.", "Chat for quick questions, mail for anything a customer or vendor sees, tickets for anything that needs an owner."]),
             ("Know where a number comes from", ["The booking system is the record for bookings and quotes, the CRM for deals.", "A spreadsheet is a copy."]), source="people/handbook/employee-handbook.md")
    deck.table("Who to ask", ["About", "Ask"], [["Pay, expenses", "Hana Kobayashi"], ["Contracts, anything to sign", "Sam Whitlock"], ["Laptops, access", "#help-it"],
                                                ["Customer data", "Isaac Brandt"], ["Anything else", "Agnes Wójcik knows who knows"]], widths=[2, 3], size=14)
    deck.save()

    # ---- Benefits 2100
    deck = Deck(path("people/benefits/benefits-2100-briefing.pptx"), "Benefits 2100", "Enrolment 26 October to 13 November 2099 · People team", owner="Dorothea Klein", modified="2099-10-26")
    deck.statement("The short version", "If you change nothing, your current choices carry over.", "Enrolment closes on Friday 13 November.")
    deck.table("What changes", ["Benefit", "2099", "2100"], [["Health plan, Earth", "Standard, Plus", "Unchanged"], ["Health plan, Luna and Mars", "Earth plan with travel cover", "New station plan with local clinics"],
                                                             ["Pension", "6% employer", "6.5% employer"], ["Commuter allowance, Earth HQ", "None", "$3,000 a month"],
                                                             ["Staff travel", "2 return trips a year on our scheduled service, standby", "3 return trips"]],
               widths=[2.2, 2.6, 2.6], source="people/benefits/benefits-2100-overview.md", size=14)
    deck.two("New in 2100", ("If you are based on Luna or Mars", ["A station health plan with clinics where you live.", "You may stay on the Earth plan with travel cover: choose during enrolment."]),
             ("If you are based at Earth HQ", ["A commuter allowance of $3,000 a month, paid with your salary.", "Nothing to do: it starts in January."]))
    deck.timeline("Dates", [("26 October", "Enrolment opens."), ("3 November", "Drop-in, 12:30, Room Deimos."), ("10 November", "Drop-in, 12:30, Room Deimos."), ("13 November", "Enrolment closes."), ("January", "Changes take effect.")])
    deck.save()

    # ---- Budget kickoff
    deck = Deck(path("finance/budget-2100/budget-2100-kickoff.pptx"), "Budget 2100", "Timetable and rules · Finance · 20 October 2099", owner="Ingrid Solberg", modified="2099-10-20")
    deck.timeline("Timetable", [("6 November", "Tomás sends each department its actuals through October."), ("20 November", "First drafts back to Finance, on the template."),
                                ("27 November", "Review meetings, 45 minutes per department."), ("10 December", "Final budget to the board.")])
    deck.bullets("Five rules", ["Start from your actuals through October.", "Headcount: every new role needs a start month and a grade.",
                                "Any single item above $250,000 needs one sentence on what happens if we do not spend it.",
                                "Do not budget for a second outer-system hull or a relief crew. Fleet options are decided separately and added centrally.",
                                "Travel between stations is budgeted by the department that travels, at the internal rate card."], source="finance/budget-2100/budget-guidelines.md")
    deck.table("The template", ["Line", "Budgeted by"], [["Revenue, scheduled service", "Finance with Operations"], ["Revenue, charter", "Sales"], ["Ship leases; traffic control charges; insurance", "Finance, centrally"], ["Fuel; ground handling and port fees; crew; maintenance", "Operations"],
                                                        ["Staff, non-crew", "Every department, with People"], ["Marketing", "Marketing"], ["Software and IT", "IT and Engineering"], ["Office, travel, other", "Every department"]],
               widths=[4, 2.5], source="finance/budget-2100/budget-template.xlsx", size=14)
    deck.table("Who approves a purchase", ["Amount in dollars", "Approval"], [["Up to 50,000", "Your manager"], ["50,000 to 250,000", "Department head"],
                                                                              ["Above 250,000", "Department head and the Finance Director"], ["Above 2,000,000", "The above and the CEO"]],
               widths=[2, 3], source="policies/purchasing-policy.md", size=14, note="A budget line is not an approval. The purchase request still comes before the order.")
    deck.save()

    # ---- Expo briefing
    deck = Deck(path("marketing/events/travel-expo-briefing.pptx"), "Interplanetary Travel Expo", "Stand briefing · Luna, 24 to 26 November 2099", owner="Pilar Navarro", modified="2099-10-28")
    deck.figures("Why we go", [("40", "qualified conversations"), ("10", "follow-up meetings"), ("1", "list of what Stellar Lines says to our customers")],
                 items=["Hall 2, stand 214. We share a wall with Stellar Lines.", "Maya Goldberg speaks on the panel \"Holding the seat: what travellers expect before they pay\", Wednesday 25 November, 14:00."],
                 source="marketing/events/travel-expo-plan.md")
    deck.table("Stand rota", ["Day", "Morning", "Afternoon"], [["Tuesday 24", "Pilar, Lena", "Pilar, Lena"], ["Wednesday 25", "Pilar, Lena", "Pilar, Maya (after the panel at 14:00)"],
                                                               ["Thursday 26", "Pilar, Jamal", "Pilar, open: one person from Sales needed"]], widths=[1.5, 2, 3], size=14)
    deck.chart("Budget: $420,000", "doughnut", ["Stand", "Everything else"], {"Dollars": [260000, 160000]}, source="marketing/events/travel-expo-plan.md and travel-expo-budget.xlsx",
               aside=("Where it goes", ["Stand: 260,000.", "Build, brochures, badges, travel and hospitality: 160,000.", "The lines are in travel-expo-budget.xlsx."]))
    deck.table("If a visitor compares us", ["They mention", "Their pitch", "Where we win", "Ask"],
               [["Orbitway", "Frequency. Twice daily Earth to Mars from March 2100.", "Pluto and Europa. A quieter cabin. Seat holds.", "How often do your crews rotate?"],
                ["Stellar Lines", "Reach. One contract for the whole system.", "Price on the inner routes, and service.", "Which destinations do you need every month?"],
                ["NovaTransit", "The lowest fare, flexible refunds.", "Reliability, business class, anything beyond Mars.", "What happens when a trip slips?"]],
               widths=[1.3, 2.6, 2.6, 2.2], source="sales/battlecards", size=12)
    deck.bullets("On the stand", ["Winter on Mars is live by then: use it. Do not say \"guaranteed aurora\".", "Never promise a departure that Operations has not confirmed.",
                                  "Write down every conversation the same day: account, contact, need, next step. It goes into the CRM on Friday.",
                                  "Brochures arrive at the Luna office on 20 November. Badges are ordered for six people."])
    deck.save()

    # ---- Pipeline review
    stage_order = ["prospect", "proposal", "negotiation"]
    by_stage = {s: [x for x in w.open_deals if x["stage"] == s] for s in stage_order}
    owners = sorted({x["owner"] for x in w.open_deals})
    soon = [x for x in w.open_deals if str(x["close_date"]) <= "2099-12-31"]
    oct_closed = [x for x in w.deals if str(x["close_date"]).startswith("2099-10") and x["stage"] in ("won", "lost")]
    pluto = [x for x in w.open_deals if x["id"] in ("D-2101", "D-2105")]
    crm = "CRM, open deals, read on Monday 2 November 2099"
    deck = Deck(path("sales/pipeline-review-2099-11-03.pptx"), "Pipeline review", "Tuesday 3 November 2099 · Sales operations", owner="Sinead Byrne", modified="2099-11-02")
    deck.figures("The open pipeline", [(str(len(w.open_deals)), "open deals"), (f"{sum(x['amount'] for x in w.open_deals) / 1e6:.1f}m", "dollars, unweighted"),
                                       (str(len(soon)), "with a close date this year"), (f"{sum(x['amount'] for x in soon) / 1e6:.1f}m", "dollars closing this year")],
                 items=["Read from the CRM, as every Tuesday. A personal tracker is a working copy, not the record."], source=crm)
    deck.chart("By stage", "column", [s.capitalize() for s in stage_order], {"Open value (dollars)": [sum(x["amount"] for x in by_stage[s]) for s in stage_order]}, source=crm,
               aside=("Deals", [f"{s.capitalize()}: {len(by_stage[s])}" for s in stage_order]))
    deck.table("By owner", ["Owner", "Open deals", "Open value (dollars)", "Closing this year"],
               [[o, sum(x["owner"] == o for x in w.open_deals), sum(x["amount"] for x in w.open_deals if x["owner"] == o), sum(x["owner"] == o for x in soon)] for o in owners],
               widths=[2, 1, 2, 1.4], source=crm, size=14)
    deck.table("Closing in November", ["Deal", "Account", "Owner", "Stage", "Close", "Value", "Next step"],
               [[x["id"], w.account(x), x["owner"].split()[0], x["stage"], short(x["close_date"]), x["amount"], x.get("next_step", "")] for x in w.open_deals if str(x["close_date"]) <= "2099-11-30"],
               widths=[0.8, 2.1, 0.9, 1.1, 0.8, 1.2, 4], source=crm, size=11)
    deck.table("Two proposals wait for the Pluto decision", ["Deal", "Account", "Close", "Value (dollars)", "Next step"],
               [[x["id"], w.account(x), short(x["close_date"]), x["amount"], x.get("next_step", "")] for x in pluto] + [["Total", "", "", sum(x["amount"] for x in pluto), ""]],
               widths=[0.8, 2.4, 0.8, 1.4, 4.2], total=True, source=crm, size=12,
               note="Both customers want two Pluto departures a month. The decision is planned for the December leadership meeting.")
    deck.table("Closed in October", ["Deal", "Account", "Owner", "Outcome", "Value (dollars)", "Loss reason"],
               [[x["id"], w.account(x), x["owner"], x["stage"].capitalize(), x["amount"], x.get("loss_reason", "")] for x in sorted(oct_closed, key=lambda x: str(x["close_date"]))],
               widths=[0.8, 2.4, 1.5, 1, 1.5, 1.8], source="CRM, deals with a close date in October 2099", size=13)
    deck.bullets("For Monday evening", ["Every open deal has a close date, an amount and a next step.", "A close date in the past on an open deal is a mistake. Fix it.",
                                        "When a deal closes, one note with the reason in the customer's own words. Won or lost."], source="sales/playbook/sales-playbook.md")
    deck.save()

    # ---- Safety pre-read
    q3 = [r for r in SAFETY if r[1] < "2099-10-01"]
    cats = Counter(r[3] for r in q3)
    deck = Deck(path("ops/safety/safety-review-2099-q3-pre-read.pptx"), "Quarterly safety review", "Third quarter 2099 · Pre-read for Tuesday 17 November", owner="Kwame Asante", modified="2099-11-03")
    deck.figures("The quarter", [(str(len(q3)), "safety reports, July to September"), ("0", "injuries"), ("0", "exceptions to the crew rest rules signed"), ("1", "bulletin issued (2099-07)")],
                 items=["Reports are never held against the person who makes them.", "Two more reports in October so far. Both are on the last slides."], source="ops/safety/safety-reports-2099.xlsx")
    deck.chart("Reports by category", "column", list(cats), {"Reports": list(cats.values())}, source="ops/safety/safety-reports-2099.xlsx",
               aside=("Reading", ["Ground and cabin, as in every quarter.", "The two Kestrel latch reports are one finding, not two."]))
    deck.table("The reports", ["No", "Date", "Where", "Category", "What"], [[r[0], short(r[1]), r[2], r[3], r[4]] for r in q3], widths=[1.2, 0.8, 1.4, 1, 5.6], size=11,
               source="ops/safety/safety-reports-2099.xlsx")
    deck.bullets("Bulletin 2099-07: cabin baggage restraints, Kestrel class", ["Two latches opened during the departure burn in September. Bags stayed in place. Nobody was hurt.",
                                                                             "Cabin crew check every latch by hand before departure, on Kestrel and Meridian.",
                                                                             "48 replacement latches are due on 14 November. Kestrel first.", "Until then, heavy items go under the seat."],
                 source="ops/safety/bulletin-2099-07.md")
    deck.two("October, for discussion", ("Halcyon, coolant valve", ["Wear found at the routine inspection after the Pluto return of 23 October. Inside limits.",
                                                                    "Assembly replaced, pressure tested for 12 hours, signed off 29 October.", "No departure moved. One week later it would have cancelled a Pluto departure."]),
             ("Tycho, assistance on arrival", ["Three travellers in October booked wheelchair assistance and nobody was there.", "The handler's assistance mailbox was not read at weekends. Changed on 20 October.",
                                               "Lucia checks the requests the day before each arrival until the end of the year."]),
             source="ops/fleet/halcyon-maintenance-2099-10.md and ticket GT-1019")
    deck.bullets("Crew rest", ["The rules were not changed and no exception was signed this year.", "A request to shorten the rest at Aphrodite to 24.5 hours from 13 December was refused on 30 October.",
                               "Either the return leg moves by 90 minutes too, or there is a second crew."], source="ops/crew/crew-rest-rules.md")
    deck.save()

    # ---- Support, October
    deck = Deck(path("support/support-review-2099-10.pptx"), "Support in October", "Monthly review · Monday 2 November 2099", owner="Nadia Haddad", modified="2099-11-02")
    deck.chart("Contacts per week", "column", [k for k, *_ in SUPPORT_WEEKS], {"Contacts in the week starting": [v for _, v, *_ in SUPPORT_WEEKS]}, source="support/weekly-numbers.md",
               aside=("Reading", ["Contacts are everything that reaches us: phone, chat and mail.", "The week of 12 October is the payment timeouts: 47 contacts in one afternoon.", "Only contacts that need follow-up become tickets."]))
    deck.table("The numbers", ["Week starting", "Contacts", "Median first reply (min)", "Satisfaction (of 5)", "Top reason"], [list(r) for r in SUPPORT_WEEKS],
               widths=[1.4, 1, 1.8, 1.6, 2.6], source="support/weekly-numbers.md", size=14)
    deck.two("Two things that were not routine", ("13 October: payment timeouts", ["45 minutes, 37 payments timed out, nobody charged twice.", "We told customers that retrying was safe. It was.",
                                                                                 "Still open: a status page we can point customers to."]),
             ("Assistance at Tycho", ["Three travellers were not met on arrival.", "Cause found and changed on 20 October. A travel credit of $100,000 for each.", "Ticket GT-1019, resolved."]))
    deck.bullets("What November brings", ["**Refund preview** in release 4.12 on 10 November. \"How much do I get back\" should fall. Greta writes the article.",
                                          "**Winter on Mars** launches the same day. Last year's spring campaign added about a fifth to contacts.",
                                          "**The later Venus departure** from 13 December: booked customers are informed this week.", "**A new colleague on Luna** starts on 1 December."])
    deck.save()

    # ---- Hold service readiness
    deck = Deck(path("engineering/architecture/hold-service-launch-readiness.pptx"), "Hold service: ready for 1 October?", "Launch readiness review · Thursday 24 September 2099",
                owner="Aditi Sharma", modified="2099-09-24")
    deck.statement("What changes for the traveller", "See the price, take fifteen minutes, then decide.", "Today the price appears at the moment of booking. From 1 October it appears first, and the seat waits.")
    deck.table("What is being launched", ["Part", "What it does", "Owner"], [["Hold service (new)", "Issues a price quote and holds the seat for 15 minutes.", "Aditi Sharma"],
                                                                            ["Booking API", "Stays the record for users, flights and bookings. Calculates the price of a booking when it is made.", "Jonas Keller"],
                                                                            ["Web frontend", "Quote page with the countdown, then payment.", "Wren Abbott"], ["Warehouse", "Nightly copy of quotes, next to bookings.", "Felix Braun"]],
               widths=[1.6, 5, 1.4], size=13)
    deck.table("Checklist", ["Check", "Owner", "Status"], [["Load test at three times the busiest hour", "Emeka Obi", "Passed 22 September"], ["Holds expire and release the seat", "Clara Jensen", "Passed"],
                                                          ["A stale hold never blocks a seat for more than 20 minutes", "Clara Jensen", "Passed"], ["Rollback: quote page off, booking as before", "Aditi Sharma", "Rehearsed 23 September"],
                                                          ["Support macro and knowledge base article", "Greta Olsen", "Ready"], ["Runbook and alerts", "Emeka Obi", "Ready"]], widths=[4.5, 1.6, 2], size=13)
    deck.bullets("Known limits at launch", ["**The quote price is calculated inside the hold service for now.** Wiring it to the booking API's calculation is the follow-up (issue 51).",
                                            "**No holds for groups.** One seat, one hold (issue 47).", "**Charter flying is not affected.** It is not in this system."])
    deck.statement("Recommendation", "Go on Thursday 1 October, for all customers.", "On call that week: Aditi, with Wren as secondary.")
    deck.save()


# ====================================================================== workbooks

def workbooks(w: World) -> None:
    d = w.d

    # ---- Departure plan, November
    rotation = {(r["origin"], r["destination"]): r for r in yaml.safe_load((COMPANY / "fleet.yaml").read_text(encoding="utf-8"))["rotations"]}
    handler = {"Earth": "Own staff, Cape Meridian", "Moon": "Tycho Ground Services", "Mars": "Olympus Port Authority", "Venus": "Aphrodite Station Operations",
               "Jupiter": "Outer Belt Port Services", "Europa": "Europa Research Trust"}
    note = {21: "Sold out.", 28: "Olympus pad 2 is closed 16 to 20 November.", 30: "Halcyon back in service since 29 October.", 26: "Last 07:00 departure. From 13 December: 08:30."}
    rows = []
    for f in w.flights.values():
        if f["departure_time"].startswith("2099-11"):
            flown, sold = rotation[f["origin"], f["destination"]], w.sold(f["flight_id"], "2099-11-02")
            rows.append({"flight": d.gx(f["flight_id"]), "route": d.route(f["flight_id"]), "departure_utc": f["departure_time"].replace("T", " ")[:16],
                         "arrival_utc": f["arrival_time"].replace("T", " ")[:16], "hours": w.hours(f), "ship": flown["ship"], "crew": flown["crew"], "seats": 10, "sold": sold,
                         "free": F("=H{r}-I{r}", 10 - sold), "handler_at_origin": handler[f["origin"]], "note": note.get(f["flight_id"], "")})
    book = Book(path("ops/schedule/departure-plan-2099-11.xlsx"), owner="Andrei Popescu", modified="2099-11-02")
    book.sheet("November departures", rows, title="Scheduled service: departure plan, November 2099",
               subtitle="Seats sold: booking system, bookings not cancelled, at Monday 2 November 2099. Ships and crews: roster of 20 October.",
               total={"seats": "sum", "sold": "sum", "free": "sum"}, bars=["sold"], widths={"note": 44, "handler_at_origin": 30})
    book.sheet("Rest rules", [{"after": "Earth to Luna shuttle (4 hours)", "minimum_rest_hours": 10}, {"after": "Inner-system leg (Mars)", "minimum_rest_hours": 16},
                              {"after": "Venus (counts as outer-system for rest)", "minimum_rest_hours": 26}, {"after": "Outer-system leg (Jupiter, Europa)", "minimum_rest_hours": 26},
                              {"after": "Pluto", "minimum_rest_hours": 26}],
               title="Minimum crew rest at destination", subtitle="From ops/crew/crew-rest-rules.md. A crew flies at most three outer-system legs in thirty days. 48 hours before the next Pluto leg.")
    book.close()

    # ---- Maintenance plan and reserve
    book = Book(path("ops/fleet/maintenance-plan-2099-q4.xlsx"), owner="Zainab Hussein", modified="2099-10-30")
    book.sheet("Plan", [
        {"ship": "Halcyon", "from": "2099-10-26", "to": "2099-10-29", "work": "Number 2 coolant valve assembly replaced, 12 hour pressure test", "kind": "Unscheduled", "where": "Cape Meridian", "status": "Done, signed off 29 October 22:10"},
        {"ship": "Kestrel", "from": "2099-11-14", "to": "2099-11-14", "work": "Baggage restraint latches replaced (bulletin 2099-07), routine check", "kind": "Scheduled", "where": "Cape Meridian", "status": "Planned. 48 latches due 14 November."},
        {"ship": "Meridian", "from": "2099-11-21", "to": "2099-11-21", "work": "Routine check. Latches after Kestrel.", "kind": "Scheduled", "where": "Cape Meridian", "status": "Planned"},
        {"ship": "Halcyon", "from": "2099-11-25", "to": "2099-11-25", "work": "Inspection after the Pluto leg of 21 November", "kind": "Scheduled", "where": "Tombaugh Station", "status": "Planned"},
        {"ship": "Solstice", "from": "2099-12-05", "to": "2099-12-05", "work": "Routine check", "kind": "Scheduled", "where": "Cape Meridian", "status": "Planned"},
        {"ship": "Aurora", "from": "2099-12-12", "to": "2099-12-12", "work": "Routine check", "kind": "Scheduled", "where": "Olympus Port", "status": "Planned"},
        {"ship": "Perihelion", "from": "2099-06-01", "to": "2100-02-28", "work": "Refit for outer-system certification", "kind": "Refit", "where": "Yard", "status": "In progress. Back in service February 2100."}],
        title="Fleet maintenance plan, fourth quarter 2099", subtitle="Kept by Zainab Hussein. Dates of the routine checks agree with ops/fleet/fleet-register.xlsx.",
        status={"status": {"done": OK, "in progress": AMBER}, "kind": {"unscheduled": WARN}}, widths={"work": 58, "status": 40})
    moves = [("2099-10-01", "Opening balance", 18500000), ("2099-10-01", "Monthly contribution", 1500000), ("2099-10-29", "Halcyon valve assembly: parts", -410000),
             ("2099-10-29", "Halcyon valve assembly: labour", -230000), ("2099-11-01", "Monthly contribution", 1500000)]
    rows, balance = [], 0
    for i, (when, what, amount) in enumerate(moves):
        balance += amount
        rows.append({"date": when, "movement": what, "amount": amount, "balance": F("=C{r}" if i == 0 else "=D" + str(i + 4) + "+C{r}", balance)})
    book.sheet("Maintenance reserve", rows, title="Maintenance reserve", subtitle="Dollars. The Halcyon repair of October cost 640,000: parts 410,000, labour 230,000.", widths={"movement": 40})
    book.close()

    # ---- Sales forecast from the CRM
    prob = {"prospect": 0.1, "proposal": 0.3, "negotiation": 0.6}
    rows = [{"deal": x["id"], "name": x["name"], "account": w.account(x), "owner": x["owner"], "stage": x["stage"], "amount": x["amount"], "close_date": str(x["close_date"]),
             "probability": prob[x["stage"]], "weighted": F("=F{r}*H{r}", round(x["amount"] * prob[x["stage"]])), "next_step": x.get("next_step", "")} for x in w.open_deals]
    book = Book(path("sales/forecast-2099-q4.xlsx"), owner="Sinead Byrne", modified="2099-11-02")
    ws, first, last = book.sheet("Open pipeline", rows, title="Open pipeline and weighted forecast", subtitle="Read from the CRM on Monday 2 November 2099. The CRM is the record: this sheet is a copy for the forecast.",
                                 total={"amount": "sum", "weighted": "sum"}, formats={"probability": "0%"}, widths={"name": 46, "next_step": 60, "account": 30},
                                 status={"stage": {"negotiation": OK, "proposal": AMBER}})
    rng = lambda c: f"'Open pipeline'!${c}${first + 1}:${c}${last + 1}"  # noqa: E731
    summary = [{"stage": s, "deals": F(f'=COUNTIF({rng("E")},A{{r}})', sum(x["stage"] == s for x in w.open_deals)),
                "open_value": F(f'=SUMIF({rng("E")},A{{r}},{rng("F")})', sum(x["amount"] for x in w.open_deals if x["stage"] == s)),
                "probability": prob[s], "weighted": F("=C{r}*D{r}", round(sum(x["amount"] for x in w.open_deals if x["stage"] == s) * prob[s]))} for s in prob]
    ws2, f2, l2 = book.sheet("By stage", summary, title="Forecast by stage", subtitle="Probabilities by stage as agreed for the fourth quarter: prospect 10%, proposal 30%, negotiation 60%.",
                             total={"deals": "sum", "open_value": "sum", "weighted": "sum"}, formats={"probability": "0%"}, widths={"open_value": 18, "weighted": 18})
    book.chart(ws2, "column", "Open value and weighted forecast by stage", "By stage", (f2, l2, 0), [("Open value", 2), ("Weighted", 4)], "G4", labels=False)
    months = sorted({str(x["close_date"])[:7] for x in w.open_deals})
    book.sheet("By close month", [{"close_month": m, "deals": sum(str(x["close_date"]).startswith(m) for x in w.open_deals),
                                   "open_value": sum(x["amount"] for x in w.open_deals if str(x["close_date"]).startswith(m)),
                                   "weighted": round(sum(x["amount"] * prob[x["stage"]] for x in w.open_deals if str(x["close_date"]).startswith(m)))} for m in months],
               title="Forecast by close month", subtitle="By the close date in the CRM.", total={"deals": "sum", "open_value": "sum", "weighted": "sum"}, bars=["open_value"])
    book.close()

    # ---- Winter on Mars media plan
    plan = [("Paid social, Earth", 110000, 90000, 80000, 60000, "Pulse Media"), ("Paid social, Luna", 70000, 60000, 50000, 40000, "Pulse Media"), ("Search", 50000, 50000, 50000, 30000, "Pulse Media"),
            ("Screens at the Tycho terminal", 60000, 60000, 60000, 0, "Tycho terminal media desk"), ("Travel press, two titles", 80000, 0, 60000, 0, "Direct"),
            ("Kepler Media films: distribution", 0, 70000, 40000, 30000, "Kepler Media Group (co-marketing)")]
    rows = [{"channel": c, "November": a, "December": b, "January": c2, "February": e, "total": F("=SUM(B{r}:E{r})", a + b + c2 + e), "booked_through": v} for c, a, b, c2, e, v in plan]
    book = Book(path("marketing/campaigns/winter-on-mars/media-plan.xlsx"), owner="Elise Fontaine", modified="2099-11-03")
    ws, first, last = book.sheet("Media plan", rows, title="Winter on Mars: paid media plan", subtitle="Dollars. 10 November 2099 to 28 February 2100. Booking closes Thursday 5 November. Not booked yet.",
                                 total={"November": "sum", "December": "sum", "January": "sum", "February": "sum", "total": "sum"}, widths={"channel": 34, "booked_through": 34})
    book.chart(ws, "column", "Spend by month", "Media plan", (first, last, 0), [("November", 1), ("December", 2), ("January", 3), ("February", 4)], "A14", width=760, labels=False, subtype="stacked")
    book.notes("Notes", ["Audience: leisure travellers from Earth and Luna who have flown with us once. Small groups.", "No discount is advertised. The public site has none; the group rate goes through Sales.",
                         "Not allowed in any copy: \"guaranteed aurora\".", "Each line above $250,000 in total needs the Finance Director's approval before the order (policies/purchasing-policy.md)."])
    book.close()

    # ---- Expo budget
    lines = [("Stand, hall 2, stand 214", 260000, 260000, "Paid"), ("Stand build and graphics", 48000, 48000, "Ordered"), ("Brochures, delivery to the Luna office 20 November", 22000, 22000, "Ordered"),
             ("Badges, six people", 24000, 24000, "Paid"), ("Station allowance on Luna, 9,000 a day", 36000, 0, "Planned"), ("Hospitality on the stand", 20000, 0, "Planned"), ("Contingency", 10000, 0, "")]
    rows = [{"line": a, "budget": b, "committed": c, "remaining": F("=B{r}-C{r}", b - c), "status": s} for a, b, c, s in lines]
    assert sum(r["budget"] for r in rows) == 420000
    book = Book(path("marketing/events/travel-expo-budget.xlsx"), owner="Pilar Navarro", modified="2099-10-28")
    ws, first, last = book.sheet("Budget", rows, title="Interplanetary Travel Expo, Luna, 24 to 26 November 2099", subtitle="Budget $420,000, of which stand 260,000. Travel between stations is at the internal rate and not in this budget.",
                                 total={"budget": "sum", "committed": "sum", "remaining": "sum"}, status={"status": {"paid": OK, "ordered": AMBER}}, widths={"line": 52})
    book.chart(ws, "pie", "Budget by line", "Budget", (first, last, 0), [("Budget", 1)], "G4", width=480)
    book.close()

    # ---- Hiring tracker
    book = Book(path("people/hiring/hiring-tracker.xlsx"), owner="Samir Chaudhry", modified="2099-11-03")
    book.sheet("Open roles", [
        {"role": "Senior Engineer, booking platform", "team": "Engineering", "location": "Earth HQ", "hiring_manager": "Marcus Lindqvist", "opened": "2099-08-24", "applicants": 64, "screened": 14, "interviewed": 6, "final_round": 2, "offers": 0, "status": "Final interviews 10 and 11 November"},
        {"role": "Support Agent", "team": "Support", "location": "Luna", "hiring_manager": "Nadia Haddad", "opened": "2099-08-31", "applicants": 41, "screened": 9, "interviewed": 4, "final_round": 2, "offers": 1, "status": "Accepted, starts 1 December"},
        {"role": "Route Analyst", "team": "Operations", "location": "Earth HQ", "hiring_manager": "Priya Raman", "opened": "2099-10-14", "applicants": 17, "screened": 3, "interviewed": 0, "final_round": 0, "offers": 0, "status": "Approved 14 October, advert live"},
        {"role": "IT Support Technician", "team": "IT", "location": "Earth HQ", "hiring_manager": "Bruno Ferreira", "opened": "2099-08-10", "applicants": 38, "screened": 8, "interviewed": 5, "final_round": 2, "offers": 1, "status": "Filled, starts 9 November"}],
        title="Hiring tracker", subtitle="At Tuesday 3 November 2099. Candidate names are in the recruiting system, not here.", total={"applicants": "sum", "screened": "sum", "interviewed": "sum", "final_round": "sum", "offers": "sum"},
        status={"status": {"accepted": OK, "filled": OK, "final": AMBER}}, widths={"role": 34, "status": 38})
    book.close()

    # ---- Support numbers
    split = [(96, 48, 27), (101, 55, 28), (139, 58, 34), (97, 52, 27), (110, 57, 31)]
    rows = []
    for (week, contacts, reply, sat, reason), (phone, chat, mail) in zip(SUPPORT_WEEKS, split):
        assert phone + chat + mail == contacts
        rows.append({"week_starting": week, "phone": phone, "chat": chat, "mail": mail, "contacts": F("=SUM(B{r}:D{r})", contacts), "median_first_reply_min": reply, "satisfaction_of_5": sat, "top_reason": reason})
    book = Book(path("support/support-numbers-2099-autumn.xlsx"), owner="Greta Olsen", modified="2099-11-03")
    ws, first, last = book.sheet("Weekly numbers", rows, title="Support: weekly numbers, autumn 2099", subtitle="Contacts: ticket system and phone log, counted by Greta Olsen each Monday. Only contacts that need follow-up become tickets.",
                                 total={"phone": "sum", "chat": "sum", "mail": "sum", "contacts": "sum"}, formats={"satisfaction_of_5": "0.0"}, widths={"top_reason": 30})
    book.chart(ws, "column", "Contacts by channel", "Weekly numbers", (first, last, 0), [("Phone", 1), ("Chat", 2), ("Mail", 3)], "A12", width=700, labels=False, subtype="stacked")
    book.close()

    # ---- Release 4.12 test matrix
    tiers = [(10, 1.0, 0.0, 0.0), (5, 0.5, 0.25, 0.25), (2, 0.0, 0.5, 0.5), (0.5, 0.0, 1.0, 0.0)]
    prices = {"economy": 1000000, "business": 2500000, "galaxium": 5000000}
    rows, n = [], 0
    for days, refund, fee, credit in tiers:
        for cls, price in prices.items():
            n += 1
            rows.append({"case": f"RP-{n:02d}", "area": "Refund preview (43)", "what": f"Cancel {days:g} days before departure", "seat_class": cls, "price_paid": price, "days_before": days,
                         "refund": F(f"=E{{r}}*{refund:g}", int(price * refund)), "fee": F(f"=E{{r}}*{fee:g}", int(price * fee)), "travel_credit": F(f"=E{{r}}*{credit:g}", int(price * credit)),
                         "status": "Passed" if n not in (11, 12) else "Open", "tester": "Clara Jensen", "note": "" if n not in (11, 12) else "Waiting for a test departure inside 24 hours"})
    blank = {"seat_class": "", "price_paid": "", "days_before": "", "refund": "", "fee": "", "travel_credit": ""}
    extra = [("RP-13", "Refund preview (43)", "Booking already cancelled: no preview", "Passed", "Clara Jensen", ""), ("RP-14", "Refund preview (43)", "Departure in the past: no preview", "Passed", "Clara Jensen", ""),
             ("CB-01", "Circuit breaker", "Provider slow: after five slow answers new payments fail fast with a clear message", "Passed", "Jonas Keller", ""),
             ("CB-02", "Circuit breaker", "Provider back: payments resume without a deploy", "Open", "Jonas Keller", "Planned for 4 November"),
             ("HT-01", "Hold timer (55)", "Slow page load: the traveller still gets the full 15 minutes", "Passed", "Wren Abbott", ""),
             ("HT-02", "Hold timer (55)", "The traveller never opens the page: the hold expires 20 minutes after the quote was created", "Open", "Clara Jensen", "Planned for 5 November")]
    rows += [{"case": c, "area": a, "what": t, **blank, "status": s, "tester": who, "note": nt} for c, a, t, s, who, nt in extra]
    book = Book(path("engineering/releases/release-4.12-test-matrix.xlsx"), owner="Clara Jensen", modified="2099-11-03")
    ws, first, last = book.sheet("Cases", rows, title="Release 4.12: test matrix", subtitle="Release Tuesday 10 November 2099, 07:00 UTC. Code freeze Friday 6 November. Refund cases on Earth to Mars fares, tiers from policies/refund-policy.md.",
                                 status={"status": {"passed": OK, "open": AMBER}}, widths={"what": 62, "note": 40, "area": 22})
    count = Counter(r["status"] for r in rows)
    book.sheet("Summary", [{"status": s, "cases": F(f"=COUNTIF(Cases!$J${first + 1}:$J${last + 1},A{{r}})", count[s])} for s in ("Passed", "Open", "Failed")],
               title="Status at 3 November", subtitle="Not in this release and not tested: group holds (47), quote price from the booking API (51).", total={"cases": "sum"}, status={"status": {"passed": OK, "open": AMBER}})
    book.close()

    # ---- Safety reports
    book = Book(path("ops/safety/safety-reports-2099.xlsx"), owner="Kwame Asante", modified="2099-10-30")
    book.sheet("Reports", [dict(zip(("report", "date", "where", "category", "what_happened", "severity", "status"), r)) for r in SAFETY], title="Safety reports, July 2099 to date",
               subtitle="Kept by the Safety Officer. Reports are never held against the person who makes them. Names of reporters are not recorded here.",
               status={"status": {"open": AMBER, "closed": OK}, "severity": {"medium": AMBER}}, widths={"what_happened": 84}, wrap=["what_happened"])
    q = Counter((r[1][:7], r[3]) for r in SAFETY)
    cats = sorted({r[3] for r in SAFETY})
    book.sheet("By month", [{"month": m, **{c: q[(m, c)] for c in cats}, "total": F(f"=SUM(B{{r}}:{chr(65 + len(cats))}{{r}})", sum(q[(m, c)] for c in cats))} for m in sorted({r[1][:7] for r in SAFETY})],
               title="Reports by month and category", subtitle="Counted from the sheet \"Reports\".", total={**{c: "sum" for c in cats}, "total": "sum"})
    book.close()

    # ---- Turnaround scorecard, third quarter
    stations = [("Earth, Cape Meridian", "Own staff", [(31, 1), (33, 0), (32, 1)]), ("Luna, Tycho", "Tycho Ground Services", [(13, 0), (15, 0), (14, 0)]),
                ("Mars, Olympus Port", "Olympus Port Authority", [(9, 0), (10, 1), (9, 0)]), ("Venus, Aphrodite Station", "Aphrodite Station Operations", [(4, 0), (4, 0), (5, 1)]),
                ("Jupiter, Ganymede Gate", "Outer Belt Port Services", [(5, 0), (5, 0), (6, 0)]), ("Europa", "Europa Research Trust", [(2, 0), (2, 0), (2, 0)]),
                ("Pluto, Tombaugh Station", "Tombaugh Station Cooperative", [(1, 0), (1, 0), (1, 0)])]
    rows = []
    for station, who, months in stations:
        for label, (handled, late) in zip(("July", "August", "September"), months):
            rows.append({"station": station, "handler": who, "month": label, "turnarounds": handled, "over_target": late, "within_target": F("=(D{r}-E{r})/D{r}", (handled - late) / handled)})
    book = Book(path("ops/ground/turnaround-scorecard-2099-q3.xlsx"), owner="Lucia Bianchi", modified="2099-10-06")
    book.sheet("Third quarter", rows, title="Turnarounds by station, third quarter 2099", subtitle="Source: operations control log. Scheduled service, charters and shuttles together. The target at every station is the one in its contract.",
               total={"turnarounds": "sum", "over_target": "sum"}, formats={"within_target": "0%"}, status={"over_target": {"1": AMBER}}, widths={"station": 28, "handler": 32})
    book.notes("Notes", ["Ground handlers count turnarounds, not departures. Tycho handles about fourteen of ours a month: one scheduled Earth to Moon, the rest shuttles and charters.",
                         "Tycho had no turnaround over 90 minutes in the quarter.", "October is in ops/ground/tycho-turnarounds-2099-10.xlsx."])
    book.close()

    # ---- Software register
    book = Book(path("it/software-register.xlsx"), owner="Bruno Ferreira", modified="2099-10-26")
    book.sheet("Software", [
        {"system": "Mail and collaboration suite", "used_for": "Mail, chat, drive, calendar", "owner": "Bruno Ferreira", "seats": len(w.people), "monthly_cost": 61000, "renews": "2099-12-15", "second_factor": "yes", "status": "Renews 15 December, automatic, 12 months"},
        {"system": "CRM", "used_for": "Accounts and deals", "owner": "Sinead Byrne", "seats": 9, "monthly_cost": 27000, "renews": "2100-04-30", "second_factor": "yes", "status": "Active"},
        {"system": "Ticket system", "used_for": "Support, IT, engineering, facilities and legal queues", "owner": "Greta Olsen", "seats": len(w.people), "monthly_cost": 19500, "renews": "2100-03-31", "second_factor": "from 9 November", "status": "Active"},
        {"system": "Payroll system", "used_for": "Payroll and payslips", "owner": "Hana Kobayashi", "seats": 4, "monthly_cost": 12000, "renews": "2100-01-31", "second_factor": "yes", "status": "Active. VPN only."},
        {"system": "Accounting ledger", "used_for": "General ledger, payables", "owner": "Chidi Eze", "seats": 6, "monthly_cost": 23000, "renews": "2100-06-30", "second_factor": "yes", "status": "Active"},
        {"system": "Warehouse", "used_for": "Nightly copies of users, flights, bookings and quotes", "owner": "Felix Braun", "seats": 11, "monthly_cost": 16000, "renews": "2100-05-31", "second_factor": "VPN", "status": "Active. Opening to Finance and Operations."},
        {"system": "Connectivity, three stations", "used_for": "Network between Earth, Luna and Mars", "owner": "Bruno Ferreira", "seats": "", "monthly_cost": 52000, "renews": "2100-02-28", "second_factor": "", "status": "Active, automatic renewal"},
        {"system": "Customer survey tool", "used_for": "Satisfaction surveys after a contact", "owner": "Nadia Haddad", "seats": 6, "monthly_cost": 14000, "renews": "", "second_factor": "", "status": "Not signed: data processing addendum open"}],
        title="Software and services register", subtitle="Kept by IT. Contract terms are in legal/vendor-register.xlsx. Second factor mandatory from 9 November 2099 for mail, chat, the drive and the CRM.",
        total={"monthly_cost": "sum"}, status={"status": {"not signed": WARN, "renews 15": AMBER}}, widths={"used_for": 48, "status": 44, "system": 30})
    book.close()

    # ---- Data requests
    reqs = [("2099-001", "2099-01-19", "Copy", 22), ("2099-002", "2099-02-03", "Copy", 17), ("2099-003", "2099-02-26", "Deletion", 28), ("2099-004", "2099-03-12", "Copy", 9), ("2099-005", "2099-04-08", "Correction", 3),
            ("2099-006", "2099-04-27", "Copy", 24), ("2099-007", "2099-05-18", "Deletion", 29), ("2099-008", "2099-06-09", "Copy", 14), ("2099-009", "2099-07-02", "Copy", 19), ("2099-010", "2099-07-30", "Correction", 5),
            ("2099-011", "2099-08-20", "Copy", 26), ("2099-012", "2099-09-11", "Deletion", 21), ("2099-013", "2099-10-05", "Copy", 16)]
    rows = [{"request": a, "received": b, "kind": k, "due": str(date.fromisoformat(b) + timedelta(days=29)), "closed_after_days": n, "status": "Closed", "note": ""} for a, b, k, n in reqs]
    rows.append({"request": "2099-014", "received": "2099-10-27", "kind": "Copy", "due": "2099-11-25", "closed_after_days": "", "status": "Open", "note": "Came in through Support on 26 October. Extract requested from Felix Braun."})
    book = Book(path("legal/privacy/data-requests-2099.xlsx"), owner="Isaac Brandt", modified="2099-10-27")
    book.sheet("Requests", rows, title="Requests from customers for their own data, 2099", subtitle="The deadline is 30 days. Travellers are identified in the ticket, not in this register.",
               total={"closed_after_days": "avg"}, total_label="Average days to close", status={"status": {"open": AMBER, "closed": OK}}, widths={"note": 60})
    book.close()

    # ---- Security training
    done = {"Legal": 3, "Finance": 4, "Support": 4, "Operations": 3, "Engineering": 3, "Sales": 2, "Marketing": 1, "IT": 1, "Leadership": 0, "People": 0}
    sizes = Counter(p["department"] for p in w.people)
    rows = [{"department": k, "people": sizes[k], "completed": v, "rate": F("=C{r}/B{r}", v / sizes[k])} for k, v in done.items()]
    assert round(100 * sum(done.values()) / len(w.people)) == 44
    book = Book(path("legal/privacy/security-training-status-2099-10-30.xlsx"), owner="Isaac Brandt", modified="2099-10-30")
    ws, first, last = book.sheet("By department", rows, title="Security and privacy training 2099: status after ten days", subtitle="At Friday 30 October 2099, 16:00. Due Friday 13 November. Source: training platform.",
                                 total={"people": "sum", "completed": "sum"}, formats={"rate": "0%"}, status={"rate": {"0.0": WARN}})
    book.chart(ws, "bar", "Completed, by department", "By department", (first, last, 0), [("Rate", 3)], "F4", number_format="0%", width=520)
    book.close()

    # ---- Purchase requests
    book = Book(path("finance/ap/purchase-requests-2099-q4.xlsx"), owner="Beatrix Novak", modified="2099-11-02")
    book.sheet("Purchase requests", [
        {"request": "PR-2099-109", "date": "2099-10-02", "requested_by": "Pilar Navarro", "department": "Marketing", "supplier": "Interplanetary Travel Expo", "what": "Stand, hall 2, stand 214", "amount": 260000, "approval_1": "Elise Fontaine", "approval_2": "Ingrid Solberg", "status": "Ordered"},
        {"request": "PR-2099-111", "date": "2099-10-06", "requested_by": "Zainab Hussein", "department": "Operations", "supplier": "Orbital Parts Supply", "what": "48 baggage restraint latches, Kestrel class", "amount": 216000, "approval_1": "Priya Raman", "approval_2": "", "status": "Ordered, delivery 14 November"},
        {"request": "PR-2099-113", "date": "2099-10-09", "requested_by": "Nadia Haddad", "department": "Support", "supplier": "Pulse Feedback", "what": "Customer survey tool, 14,000 a month", "amount": 168000, "approval_1": "Olu Adeyemi", "approval_2": "", "status": "On hold: data processing addendum not signed"},
        {"request": "PR-2099-115", "date": "2099-10-15", "requested_by": "Andrei Popescu", "department": "Operations", "supplier": "Brightwell Training", "what": "Simulator sessions, fourth quarter", "amount": 180000, "approval_1": "Priya Raman", "approval_2": "", "status": "Ordered"},
        {"request": "PR-2099-117", "date": "2099-10-24", "requested_by": "Zainab Hussein", "department": "Operations", "supplier": "Orbital Parts Supply", "what": "Coolant valve assembly, Halcyon", "amount": 410000, "approval_1": "Priya Raman", "approval_2": "Ingrid Solberg", "status": "Delivered, paid from the maintenance reserve"},
        {"request": "PR-2099-118", "date": "2099-10-23", "requested_by": "Bruno Ferreira", "department": "IT", "supplier": "Laptop supplier", "what": "Two replacement laptops (Andrei Popescu, Hana Kobayashi)", "amount": 158000, "approval_1": "Bruno Ferreira", "approval_2": "", "status": "Ordered 26 October, delivery 5 November"},
        {"request": "PR-2099-120", "date": "2099-10-28", "requested_by": "Pilar Navarro", "department": "Marketing", "supplier": "Print shop", "what": "Expo brochures", "amount": 22000, "approval_1": "Elise Fontaine", "approval_2": "", "status": "Ordered"},
        {"request": "PR-2099-122", "date": "2099-11-02", "requested_by": "Elise Fontaine", "department": "Marketing", "supplier": "Pulse Media", "what": "Winter on Mars paid media, November to February", "amount": sum(a + b + c + e for _, a, b, c, e, _ in plan), "approval_1": "Elise Fontaine", "approval_2": "", "status": "Waiting for the Finance Director"}],
        title="Purchase requests, fourth quarter 2099", subtitle="Above $250,000: department head and the Finance Director. A request comes before the order, not after the invoice.",
        total={"amount": "sum"}, status={"status": {"on hold": WARN, "waiting": AMBER, "delivered": OK}}, widths={"what": 50, "status": 44, "supplier": 28})
    book.close()


# ====================================================================== documents

def documents(w: World) -> None:
    # ---- The current Tycho agreement, behind legal/vendors/tycho/current-agreement-summary.md
    doc = Doc(path("legal/vendors/tycho/ground-handling-agreement-2098.docx"), "Ground Handling Agreement", subtitle="Tycho Spaceport, Luna", brand=AGREEMENT,
              label="Signed 12 December 2097 · Contract no. V-2097-031", owner="Sam Whitlock", modified="2097-12-12")
    doc.p("This Agreement is made on 12 December 2097 between **Galaxium Travels** (\"Customer\"), Cape Meridian, Earth, and **Tycho Ground Services Ltd.** (\"Vendor\"), Tycho Spaceport, Luna.", align="j")
    for number, title, paragraphs in [
        ("1", "Services", ["The Vendor provides ground handling for the Customer's departures and arrivals at Tycho Spaceport: docking, fuelling, baggage and passenger transfer.",
                           "The Vendor provides assistance to travellers who have requested it. The Customer sends each request to the Vendor's assistance desk no later than 48 hours before arrival."]),
        ("2", "Term", ["This Agreement runs for 24 months, from 1 January 2098 to 31 December 2099.", "It does not renew automatically. A renewal requires a new agreement signed by both parties."]),
        ("3", "Fees", ["The Customer pays a monthly fee of $380,000 for up to 16 turnarounds per month. Each further turnaround is charged at $30,000."]),
        ("4", "Price adjustment", ["The fees are fixed for the term."]),
        ("5", "Payment", ["The Vendor invoices monthly in arrears and attaches the monthly service report to each invoice.", "Invoices are payable within 45 days of receipt of the invoice and the service report."]),
        ("6", "Service levels", ["The Vendor completes each turnaround within 90 minutes of docking.",
                                 "For each turnaround that takes longer than 90 minutes the Customer is entitled to a service credit of 5% of the monthly fee.",
                                 "The Customer claims a service credit in writing within 30 days of the turnaround concerned. Credits are deducted from the next invoice."]),
        ("7", "Confidentiality", ["Each party keeps the other party's confidential information secret for the term and for three years after it ends."]),
        ("8", "Limitation of liability", ["The total liability of each party under this Agreement in any contract year is limited to the fees paid in the twelve months before the event giving rise to the claim."]),
        ("9", "Termination", ["Either party may terminate this Agreement for material breach that is not remedied within 30 days of written notice.", "The Customer may terminate this Agreement for convenience with 60 days' written notice."]),
        ("10", "Incident notification", ["The Vendor notifies the Customer of any safety or data incident affecting the Customer's passengers or operations within 24 hours of becoming aware of it."]),
        ("11", "Governing law", ["This Agreement is governed by the Earth Commercial Code."]),
        ("12", "Entire agreement", ["This Agreement is the entire agreement between the parties on its subject and replaces all earlier arrangements."])]:
        doc.clause(number, title, paragraphs)
    doc.signatures([("For Galaxium Travels", "Olu Adeyemi", "Chief Operating Officer", "Signed 12 December 2097"), ("For Tycho Ground Services Ltd.", "Rafael Ortega", "Account Director", "Signed 12 December 2097")])
    doc.save()

    # ---- Tycho's service report for September, the document that was missing from the October invoice
    turns = [(1, "Luna shuttle", 71), (3, "Luna shuttle", 68), (5, "Scheduled service", 84), (8, "Charter", 79), (10, "Luna shuttle", 66), (12, "Luna shuttle", 73), (15, "Scheduled service", 81),
             (17, "Charter", 88), (19, "Luna shuttle", 70), (22, "Luna shuttle", 64), (24, "Charter", 77), (26, "Luna shuttle", 69), (28, "Luna shuttle", 75), (30, "Charter", 82)]
    doc = Doc(path("finance/ap/tycho-service-report-2099-09.docx"), "Monthly service report", subtitle="Galaxium Travels · September 2099", brand=TYCHO, label="Service report · Attached to invoice TG-2099-09",
              meta=[("Customer", "Galaxium Travels"), ("Period", "1 to 30 September 2099"), ("Agreement", "Ground Handling Agreement of 12 December 2097"), ("Prepared by", "Rafael Ortega, 1 October 2099")],
              owner="Beatrix Novak", modified="2099-10-02")
    doc.h1("Summary")
    doc.table(["Measure", "September"], [["Turnarounds handled", len(turns)], ["Completed within 90 minutes", len(turns)], ["Longer than 90 minutes", 0], ["Longest turnaround (minutes)", max(t[2] for t in turns)],
                                         ["Assistance requests received and met", 6], ["Safety or data incidents notified", 0]], widths=[4, 1.4])
    doc.h1("Turnarounds")
    doc.table(["Date", "Flight", "Minutes from docking", "Target", "Within target"], [[f"{day} September", kind, minutes, 90, "Yes"] for day, kind, minutes in turns], widths=[1.6, 2, 1.6, 0.9, 1.2])
    doc.h1("Remarks")
    doc.p("No service credit arises under clause 6 for September. From October the fuelling gantry at stand 4 is shared with a second operator on some days. A second gantry is planned.")
    doc.save()

    # ---- Lagrange Catering, signed 21 October
    doc = Doc(path("legal/vendors/lagrange-catering-agreement-2099.docx"), "Cabin Catering Agreement", subtitle="Lagrange Catering", brand=AGREEMENT,
              label="Signed 21 October 2099 · Contract no. V-2099-044", owner="Sam Whitlock", modified="2099-10-21")
    doc.p("This Agreement is made on 21 October 2099 between **Galaxium Travels** (\"Customer\") and **Lagrange Catering** (\"Vendor\").", align="j")
    for number, title, paragraphs in [
        ("1", "Services", ["The Vendor supplies cabin catering for the Customer's departures from Cape Meridian and Tycho, to the menus and quantities ordered by the Customer for each departure."]),
        ("2", "Term", ["This Agreement runs for 12 months, from 21 October 2099 to 20 October 2100.", "It does not renew automatically."]),
        ("3", "Prices", ["Prices are those of the price list in Annex 1 and are fixed for the term.", "The Vendor invoices by delivery. The Customer pays for the quantities delivered, as shown on the signed delivery note."]),
        ("4", "Payment", ["Invoices are payable within 45 days."]),
        ("5", "Service levels", ["Each delivery is on board no later than 60 minutes before departure. For each late or incomplete delivery the Customer is entitled to a credit of 5% of that month's invoices."]),
        ("6", "Food safety and incidents", ["The Vendor notifies the Customer of any food safety incident affecting a delivery within 24 hours of becoming aware of it."]),
        ("7", "Limitation of liability", ["The total liability of each party in any contract year is limited to the fees paid in the twelve months before the event giving rise to the claim."]),
        ("8", "Termination", ["Either party may terminate for material breach not remedied within 30 days of written notice.", "The Customer may terminate for convenience with 60 days' written notice."]),
        ("9", "Governing law", ["This Agreement is governed by the law of the Earth Common Jurisdiction."])]:
        doc.clause(number, title, paragraphs)
    doc.signatures([("For Galaxium Travels", "Olu Adeyemi", "Chief Operating Officer", "Signed 21 October 2099"), ("For Lagrange Catering", "Mireille Aubin", "Managing Director", "Signed 21 October 2099")])
    doc.save()

    # ---- Charter order, Vesta Shipwrights (D-2019)
    vesta = next(x for x in w.deals if x["id"] == "D-2019")
    quarter = vesta["amount"] // 4
    doc = Doc(path("sales/contracts/vesta-shipwrights-charter-order-2100.docx"), "Charter order", subtitle="Vesta Shipwrights · Yard crew rotation 2100", label="Customer contract · Confidential",
              meta=[("Customer", "Vesta Shipwrights, Asteroid Belt"), ("Customer contact", "Tariq Hassan, Yard Manager"), ("Account executive", "Dev Malhotra"), ("Deal", "D-2019"),
                    ("Signed", long(vesta["close_date"])), ("Term", "1 January to 31 December 2100")], owner="Dev Malhotra", modified=str(vesta["close_date"]))
    doc.h1("1. What we fly")
    doc.p("One crew rotation a month for the yard, on dates fixed for the year in Annex A. The yard works a monthly cycle, so one departure a month covers every crew change.")
    doc.h1("2. Price and payment")
    doc.table(["Item", "Dollars"], [["Contract value for twelve months, after discount", vesta["amount"]], ["Payable by quarter in advance, four instalments of", quarter]], widths=[4, 1.5])
    doc.p("The price is the rate card price less 8%. The discount was approved in the deal desk on condition of payment by quarter in advance (policies/discount-policy.md). It applies to the base fare, not to fuel surcharges or port fees.")
    doc.h1("3. Names and changes")
    doc.bullets(["Names are supplied at booking and must match the travel document.", "A change of traveller is a cancellation and a new booking. Inside the contract it carries no fee up to 72 hours before departure.",
                 "Cancelled seats are returned as travel credit valid for twelve months, not as cash."])
    doc.h1("4. Schedule changes")
    doc.p("Galaxium gives 14 days' written notice of a change to a departure in Annex A. A departure is confirmed only once Operations has confirmed it.")
    doc.h1("5. General terms")
    doc.p("The Galaxium master terms for corporate customers apply. Liability is limited to the fees for the affected bookings.")
    doc.signatures([("For Galaxium Travels", "Amara Okafor", "Head of Sales", f"Signed {long(vesta['close_date'])}"), ("For Vesta Shipwrights", "Tariq Hassan", "Yard Manager", f"Signed {long(vesta['close_date'])}")])
    doc.save()

    # ---- Case study
    doc = Doc(path("sales/case-studies/helix-mining-case-study.docx"), "A shift calendar that flies", subtitle="Helix Mining Cooperative and Galaxium Travels", label="Customer story · Approved for external use",
              meta=[("Customer", "Helix Mining Cooperative, Mars"), ("Segment", "Mining and Resources"), ("Customer since", "2096"), ("What we fly", "Crew rotation charter")],
              owner="Jamal Carter", modified="2099-10-14")
    doc.h1("The need")
    doc.p("Helix runs its sites on fixed shifts. A crew that arrives a day late is a shift that does not start. What Helix asks of a carrier is not the lowest fare. It is that the ship leaves on the day in the calendar.")
    doc.h1("What we built")
    doc.bullets(["A rotation charter on dates fixed a year ahead, agreed with Operations before the contract was signed.", "Whole-ship departures: ten seats, one crew.",
                 "One contact for changes: the customer success manager, with a quarterly review."])
    doc.callout("\"We plan shifts a year ahead. We need a carrier that plans the same way.\"  Oksana Petrenko, Travel Manager, Helix Mining Cooperative", "In the customer's words")
    doc.h1("What we are changing next")
    doc.p("At the October review Helix asked for invoices by site instead of by traveller. That starts with the January invoice.")
    doc.h1("About Galaxium Travels")
    doc.p("Six ships, ten scheduled routes from Earth to Pluto, and charter rotations built around the customer's calendar. sales@galaxium.example")
    doc.save()

    # ---- Job description and scorecard
    doc = Doc(path("people/hiring/job-description-route-analyst.docx"), "Route Analyst", subtitle="Flight Operations · Earth HQ, Cape Meridian", label="Job description",
              meta=[("Reports to", "Priya Raman, Head of Flight Operations"), ("Works with", "Yuki Tanaka, Route Planner"), ("Approved", "14 October 2099"), ("Recruiter", "Samir Chaudhry")],
              owner="Samir Chaudhry", modified="2099-10-14")
    doc.h1("Why this role exists")
    doc.p("Every Monday the route review decides where our ships fly. Today one route planner prepares it, from numbers pulled by hand. We want a second person who makes those numbers quick, checked and readable.")
    doc.h1("What you will do")
    doc.bullets(["Prepare the weekly route review: seats sold per departure from the booking system, charter demand from the CRM. They are different sources and do not add up to one number.",
                 "Work with the data engineer so that the weekly numbers come from the warehouse and not from a spreadsheet.", "Write the papers behind schedule decisions: one question, the options, what each needs, what it costs.",
                 "Say where every figure comes from, and say when you are not sure."])
    doc.h1("What you bring")
    doc.bullets(["Three years or more of analysis in transport, logistics or a similar scheduled business.", "Comfortable with queries and with a spreadsheet, and clear about which of the two is the copy.",
                 "Plain writing. One page beats twenty slides."])
    doc.h1("What we offer")
    doc.bullets(["25 days off a year, plus the station holidays.", "Pension with 6.5% from the employer from 2100.", "Three return trips a year on our scheduled service, standby."])
    doc.save()

    doc = Doc(path("people/hiring/interview-scorecard-senior-engineer.docx"), "Interview scorecard", subtitle="Senior Engineer, booking platform · Final round", label="People · Confidential",
              meta=[("Candidate", "A / B (circle)"), ("Interviewer", ""), ("Date", "10 or 11 November 2099"), ("Session", "System design (Aditi) · Team and product (Marcus, Maya)")],
              owner="Samir Chaudhry", modified="2099-11-03")
    doc.p("Fill this in on the day, before you talk to the other interviewers. Score each line from 1 (no evidence) to 4 (strong evidence), and write what you saw, not what you felt.")
    doc.table(["What we look for", "Score 1 to 4", "Evidence"], [["Designs for failure: timeouts, retries, a slow provider", "", ""], ["Keeps one source of truth: no second calculation of the same thing", "", ""],
                                                               ["Explains a trade-off to a non-engineer", "", ""], ["Reviews and is reviewed without heat", "", ""], ["Writes things down", "", ""]], widths=[3.2, 1, 3])
    doc.h2("Would you hire?")
    doc.p("Strong yes  ·  Yes  ·  No  ·  Strong no")
    doc.h2("One thing you would want to know before deciding")
    doc.p(" ")
    doc.save()

    # ---- Offer letter template
    doc = Doc(path("people/templates/offer-letter-template.docx"), "Offer of employment", label="People · Template", owner="Dorothea Klein", modified="2099-09-01")
    doc.p("[Date]")
    doc.p("Dear [first name],")
    doc.p("We are glad to offer you the position of **[role]** in our **[team]** team at Galaxium Travels, based at **[Earth HQ, Cape Meridian / the Luna Station office / the Mars office]**, reporting to **[manager]**.")
    doc.table(["Term", ""], [["Start date", "[date]"], ["Salary", "$[amount] a year, paid monthly on the 28th"], ["Working time", "Full time. Core hours 10:00 to 15:00 in your station's time."],
                             ["Time off", "25 days a year, plus the station holidays where you are based"], ["Pension", "[6% / 6.5% from 2100] from the employer"],
                             ["Staff travel", "[2 / 3 from 2100] return trips a year on our scheduled service, standby"], ["Probation", "Three months"]], widths=[1.4, 4])
    doc.p("This offer is subject to the usual checks. The employee handbook and the code of conduct are part of your contract; you will find both in the drive on your first day.")
    doc.p("Please confirm by replying to this letter by [date]. If you have questions, call me.")
    doc.p("Kind regards,")
    doc.signatures([("For Galaxium Travels", "Dorothea Klein", "Head of People", ""), ("Accepted", "[name]", "", "")])
    doc.save()

    # ---- Station guides
    def here(place: str) -> list[list[str]]:
        return [[p["name"], p["role"], p["department"]] for p in w.at(place)]

    doc = Doc(path("people/handbook/station-guide-luna.docx"), "Station guide: Luna office", subtitle="Tycho terminal, level 3", label="People · Station guide",
              meta=[("Owner", "Dorothea Klein"), ("Local contact", "Lucia Bianchi"), ("Updated", "20 October 2099"), ("Allowance", "$9,000 a day from 1 November")], owner="Dorothea Klein", modified="2099-10-20")
    doc.h1("Getting there")
    doc.p("The office is on level 3 of the Tycho terminal, above the crew lounge. Badge access from 06:00 to 22:00. Staff travel on the Luna shuttle in economy; book through the internal travel form, not the public site.")
    doc.h1("Working hours")
    doc.p("Core hours are 10:00 to 15:00 Luna time. Earth HQ is two hours ahead.")
    doc.h1("Lunch")
    doc.p("There is no canteen. The lunch allowance is paid with your salary.")
    doc.h1("Who is here")
    doc.table(["Name", "Role", "Team"], here("Luna Station office"), widths=[1.6, 2.6, 1.2])
    doc.h1("Good to know")
    doc.bullets(["Ground handling at Tycho is done by Tycho Ground Services. Lucia is our contact for anything on the apron.", "Stays of more than 14 days need approval before travel.",
                 "Health: from 2100 the station plan covers the clinic in the terminal, level 1."])
    doc.save()

    doc = Doc(path("people/handbook/station-guide-mars.docx"), "Station guide: Mars office", subtitle="Olympus Port", label="People · Station guide",
              meta=[("Owner", "Dorothea Klein"), ("Local contact", "Hiro Matsuda"), ("Updated", "20 October 2099"), ("Allowance", "$11,000 a day from 1 November")], owner="Dorothea Klein", modified="2099-10-20")
    doc.h1("Getting there")
    doc.p("The office is in the operations building at Olympus Port, next to the Galaxium desk in the arrivals hall. Staff travel on our scheduled Earth to Mars service in economy.")
    doc.h1("Working hours")
    doc.p("Core hours are 10:00 to 15:00 Mars office time. A Mars day is 39 minutes longer than an Earth day: the shared calendar shows meetings in your own time, so trust the calendar, not your arithmetic.")
    doc.h1("The Galaxium desk")
    doc.p("Open from 06:00 to 22:00 local time. Lost property from our flights is kept here; travellers collect it with a photo ID.")
    doc.h1("Who is here")
    doc.table(["Name", "Role", "Team"], here("Mars office"), widths=[1.6, 2.6, 1.2])
    doc.h1("Good to know")
    doc.bullets(["Ground handling is done by the Olympus Port Authority. Notices to operators go to operations control.", "Stays of more than 14 days need approval before travel.",
                 "Health: from 2100 the station plan covers the port clinic."])
    doc.save()

    # ---- Data processing addendum
    doc = Doc(path("legal/templates/data-processing-addendum.docx"), "Data processing addendum", subtitle="Template for vendors that process personal data for Galaxium Travels", label="Legal · Template · Version 3",
              meta=[("Owner", "Isaac Brandt"), ("Approved by", "Helena Vogt"), ("Version", "3, 1 September 2099"), ("Use", "Signed before the vendor receives any data")], owner="Isaac Brandt", modified="2099-09-01")
    doc.callout("Clauses 3 and 5 are not negotiable without the General Counsel: notice of sub-processors before the change, and deletion within 30 days.", "For reviewers")
    for number, title, paragraphs in [
        ("1", "Scope", ["This addendum applies whenever the Vendor processes personal data of Galaxium's customers or staff. For customers this means travel data: names, itineraries, contact details and payment status."]),
        ("2", "Instructions", ["The Vendor processes personal data only on Galaxium's documented instructions and only for the service described in the main agreement."]),
        ("3", "Sub-processors", ["The Vendor informs Galaxium at least 30 days before engaging a new sub-processor.", "Galaxium may object. If the parties do not agree, Galaxium may terminate the main agreement without penalty."]),
        ("4", "Security", ["The Vendor keeps appropriate technical and organisational measures.", "The Vendor reports any incident affecting Galaxium's data within 24 hours of becoming aware of it."]),
        ("5", "Deletion", ["At the end of the contract the Vendor deletes all personal data within 30 days and confirms this in writing."]),
        ("6", "Requests from individuals", ["The Vendor forwards any request from an individual about their data to legal@galaxium.example within two working days and does not answer it itself."]),
        ("7", "Audit", ["Galaxium may audit the Vendor once a year with 30 days' notice."])]:
        doc.clause(number, title, paragraphs)
    doc.signatures([("For Galaxium Travels", "", "", ""), ("For the Vendor", "", "", "")])
    doc.save()

    # ---- Audit engagement letter
    doc = Doc(path("finance/audit/engagement-letter-2099.docx"), "Engagement letter", subtitle="Audit of the financial statements for the year ending 31 December 2099", brand=MARLOWE, label="Private and confidential",
              owner="Chidi Eze", modified="2099-09-16")
    doc.p("14 September 2099")
    doc.p("Ingrid Solberg\nFinance Director, Galaxium Travels\nCape Meridian, Earth")
    doc.p("Dear Ms Solberg,")
    doc.p("Thank you for asking us to audit the financial statements of Galaxium Travels for the year ending 31 December 2099. This letter sets out the scope of our work, the timetable and our fee.", align="j")
    doc.h2("Scope")
    doc.p("We will audit the financial statements and report to the board whether they give a true and fair view. This year we will pay particular attention to:", align="j")
    doc.bullets(["revenue from the scheduled service, including the price quotes that you plan to introduce on 1 October;", "deferred revenue for bookings that depart after the year end;",
                 "the three largest vendor contracts and the payables that arise from them;", "refunds paid, and the policy tier applied to each;", "payroll."])
    doc.h2("Timetable")
    doc.table(["Step", "When"], [["First request list", "By 23 October 2099"], ["Interim fieldwork at Cape Meridian", "Week of 23 November 2099"], ["Year-end fieldwork", "February 2100"], ["Report to the board", "March 2100"]], widths=[3, 2])
    doc.h2("What we need from you")
    doc.p("Access to the general ledger, the booking system's records of bookings and quotes, and the people who prepare the monthly close. Where a figure is taken from an export, we will ask to see it agreed to the system it came from.", align="j")
    doc.h2("Fee")
    doc.p("Our fee for the audit is $1,480,000, invoiced in three parts: on signature, after the interim fieldwork and on our report.", align="j")
    doc.p("Yours sincerely,")
    doc.signatures([("For Marlowe & Finch", "Oswin Marlowe", "Partner", "Signed 14 September 2099"), ("Agreed for Galaxium Travels", "Ingrid Solberg", "Finance Director", "Signed 16 September 2099")])
    doc.save()

    # ---- Press release
    doc = Doc(path("marketing/press/press-release-seat-holds-2099-10-01.docx"), "Galaxium Travels introduces seat holds", subtitle="See the price, take fifteen minutes, then decide", label="Press release · For release 1 October 2099, 09:00",
              owner="Noor Rahimi", modified="2099-09-30")
    doc.p("**Cape Meridian, 1 October 2099.** From today, travellers on all ten scheduled routes of Galaxium Travels see the price of a seat before they book it, and the seat is held for them for fifteen minutes while they decide.")
    doc.p("Until now the price appeared at the moment of booking. With the new hold service a traveller chooses a departure and a seat class, is shown a price, and has fifteen minutes to pay. If the hold runs out, the seat goes back on sale and nothing is charged.")
    doc.callout("\"A departure with ten seats fills quickly. People should not have to choose between thinking and losing the seat.\"  Renata Castellanos, Chief Executive Officer", "Quote")
    doc.h2("How it works")
    doc.bullets(["Choose a departure and a seat class: economy, business or galaxium.", "We show the price and hold the seat for fifteen minutes.", "Pay inside the hold and the seat is booked. Otherwise the hold ends by itself."], numbered=True)
    doc.h2("About Galaxium Travels")
    doc.p("Galaxium Travels flies scheduled service on ten routes between Earth, Luna, Mars, Venus, Jupiter, Europa and Pluto, and charter rotations for corporate customers. Six ships, ten seats each.")
    doc.h2("Contact")
    doc.p("Noor Rahimi, Content Lead · press@galaxium.example")
    doc.save()

    # ---- Hold service design
    doc = Doc(path("engineering/architecture/hold-service-design.docx"), "Hold service: design", subtitle="Price quotes and seat holds for the scheduled service", label="Engineering · Design document",
              meta=[("Author", "Aditi Sharma"), ("Reviewers", "Marcus Lindqvist, Jonas Keller, Maya Goldberg"), ("Status", "Approved 11 September 2099"), ("Goes live", "1 October 2099")],
              owner="Aditi Sharma", modified="2099-09-18")
    doc.h1("1. What it is for")
    doc.p("A traveller should see the price before paying and should not lose the seat while reading it. The hold service issues a price quote and holds one seat for 15 minutes. It does not make bookings: the booking API stays the system of record for users, flights and bookings.")
    doc.h1("2. The flow")
    doc.bullets(["The web frontend asks the hold service for a quote: user, flight, seat class.", "The hold service checks that a seat of that class is free, holds it and answers with a quote and its price.",
                 "The traveller pays inside the hold. The frontend calls the booking API, which makes the booking and calculates the price of the booking.", "The hold service marks the quote as converted and stores the booking id.",
                 "If nothing happens for 15 minutes, the expiry job marks the quote as expired and releases the seat."], numbered=True)
    doc.h1("3. What a quote is")
    doc.table(["Field", "Meaning"], [["quote_id", "Number of the quote"], ["user_id, flight_id, seat_class", "Who asked, for which departure and class"], ["quoted_price", "The price shown to the traveller, in dollars"],
                                     ["created_at", "When the quote was issued (UTC). The hold runs from here."], ["status", "converted or expired"], ["booking_id", "The booking it became, empty for an expired quote"]], widths=[2, 4])
    doc.p("Quotes are copied to the warehouse every night, next to bookings.")
    doc.h1("4. Rules")
    doc.bullets(["One quote holds one seat. There are no holds for groups in this version.", "A hold lasts 15 minutes. A hold older than 20 minutes is a bug.", "A quote is not a booking, and an expired quote charges nothing."])
    doc.h1("5. Price")
    doc.callout("For the launch the hold service calculates the quote price itself, from a simple table. The booking API calculates the price of the booking when it is made. Wiring the quote to the booking API's calculation is the follow-up, issue 51.", "Interim")
    doc.p("The pricing policy says that the fare shown on a quote must equal the fare charged at booking.")
    doc.h1("6. Operation")
    doc.bullets(["Alerts on hold duration and on the age of the oldest hold.", "Rollback: the quote page can be switched off, and booking works as before.", "On call: the booking platform rota."])
    doc.save()

    # ---- Turnaround procedure, Tycho
    steps = [("Docking and safing", "Handler", 8), ("Disembarkation, including assistance", "Cabin crew, handler", 12), ("Fuelling at the gantry", "Handler", 35),
             ("Cabin check: every restraint latch by hand on Kestrel class", "Cabin crew", 10), ("Boarding", "Cabin crew, handler", 20), ("Release and undocking", "Operations control, handler", 5)]
    assert sum(s[2] for s in steps) == 90
    doc = Doc(path("ops/ground/turnaround-procedure-tycho.docx"), "Turnaround procedure: Tycho", subtitle="Ground operations manual, section 4", label="Operations · Procedure",
              meta=[("Owner", "Lucia Bianchi"), ("Version", "3, 20 October 2099"), ("Handler", "Tycho Ground Services"), ("Target", "90 minutes from docking")], owner="Lucia Bianchi", modified="2099-10-20")
    doc.h1("1. The target")
    doc.p("A turnaround at Tycho is complete within 90 minutes of docking. This is the service level in our agreement with the handler, and the roster is built on it.")
    doc.h1("2. The time budget")
    doc.table(["Step", "Who", "Minutes"], [list(s) for s in steps] + [["Total", "", 90]], widths=[4, 2.2, 1], total=True)
    doc.p("Baggage, catering and cabin cleaning run alongside fuelling and are not on the critical path. Fuelling is: if the gantry is not free at docking, tell operations control at once.")
    doc.h1("3. Assistance on arrival")
    doc.bullets(["Support sends each assistance request to assistance@tychoground.example when it is booked. The mailbox is read daily, weekends included, since 20 October 2099.",
                 "The day before each arrival the ground operations manager checks the handler's list against our bookings.", "If nobody is at the door, cabin crew assist and report it in #support-escalations, naming the station."])
    doc.h1("4. When a turnaround runs over")
    doc.bullets(["Operations control writes the minutes and the cause in the control log.", "The ground operations manager tells Legal. A service credit must be claimed in writing within 30 days.", "A late turnaround is never recovered by shortening the cabin check."])
    doc.callout("Anyone may stop a turnaround for a safety reason.", "Safety first")
    doc.save()

    # ---- Board minutes, August
    doc = Doc(path("leadership/board/board-minutes-2099-08-20.docx"), "Minutes of the board meeting", subtitle="Thursday 20 August 2099, 09:00 to 12:10 · Boardroom, Earth HQ", label="Board · Confidential",
              meta=[("Chair", "Margit Halvorsen"), ("Directors", "Desmond Achebe, Katarzyna Lis, Renata Castellanos"), ("In attendance", "Ingrid Solberg, Olu Adeyemi, Farid Nazari"), ("Secretary", "Helena Vogt")],
              owner="Helena Vogt", modified="2099-08-27")
    doc.h1("1. Minutes and matters arising")
    doc.p("The minutes of the meeting of 21 May 2099 were approved. No matters arising.")
    doc.h1("2. Chief Executive's report")
    doc.p("The scheduled service flies ten routes with one departure a month each. The Pluto route, added in the spring, fills steadily. The hold service, which shows travellers a price and holds the seat for 15 minutes, is on track for 1 October.")
    doc.h1("3. Finance")
    doc.p("The Finance Director presented bookings for the first half of August, cash and the payables ageing at 31 July. Ms Lis asked which system the bookings table came from, when it was pulled and whether it agreed to the ledger. The answer could not be given in the meeting.")
    doc.callout("From the next pack on, every table states the system or file it comes from, the date it was pulled and who pulled it.", "The board asked")
    doc.h1("4. Fleet")
    doc.p("Five ships are in service. Perihelion is in refit until February 2100 and will be the second outer-system hull. Until then Halcyon flies Europa and Pluto alone. Mr Achebe asked what happens if Halcyon is out of service; the Chief Operating Officer answered that an outer-system departure would be cancelled, and that this is the largest operational risk of the winter.")
    doc.h1("5. Risks")
    doc.bullets(["One outer-system hull until February 2100.", "Two price calculations once the hold service is live. The Chief Technology Officer intends to merge them.", "Finance and Operations rely on exports pulled by hand."])
    doc.h1("6. Resolutions")
    doc.p("None.")
    doc.h1("7. Actions")
    doc.table(["Action", "Owner", "For"], [["Sources on every table of the board pack", "Ingrid Solberg", "19 November"], ["Recommendation on the Luna ground handling agreement, which ends 31 December", "Helena Vogt", "19 November"],
                                           ["2100 budget envelope", "Ingrid Solberg", "19 November"], ["Fleet options for 2100", "Olu Adeyemi", "December"]], widths=[4, 1.6, 1.2])
    doc.p("Next meeting: Thursday 19 November 2099, 09:00.")
    doc.signatures([("Chair", "Margit Halvorsen", "", "Approved 27 August 2099"), ("Secretary", "Helena Vogt", "General Counsel", "27 August 2099")])
    doc.save()


def build(d) -> None:
    w = World(d)
    decks(w)
    workbooks(w)
    documents(w)
    worldbook.build(d)
    companybook.build(d)
    print(f"documents: {sum(1 for _ in DRIVE.rglob('*.pptx'))} decks, {sum(1 for _ in DRIVE.rglob('*.docx'))} documents, {sum(1 for _ in DRIVE.rglob('*.xlsx'))} workbooks in the generated drive")
