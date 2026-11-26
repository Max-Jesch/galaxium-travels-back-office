"""The ships, the ports and the cabin: what the company knows about its own hardware and
the places it flies to.

`documents.build()` calls `build()`. Everything lands in `content/generated/drive/`:

    ops/fleet/      the fleet guide (deck), a data sheet per class, the specifications workbook
    ops/ground/     a brief per port, the network deck, the port directory
    marketing/      the cabin class guide
    support/        what travellers are told before they fly

The facts are in three files under `content/company/`: `fleet.yaml`, `ports.yaml` and
`cabin.yaml`. Nothing about a ship, a port or a seat class is typed in here. Routes and
block hours are read from the snapshot. The rules of `documents.py` apply: figures from
existing notes are repeated exactly, nothing answers a question that a pack or an open
mail asks, and every document is dated on or before 4 November 2099.

`check()` fails the build if the three files stop agreeing with the snapshot or with
each other.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import yaml

from office import ORANGE, Book, Deck, Doc

ROOT = Path(__file__).resolve().parent.parent
DRIVE = ROOT / "content" / "generated" / "drive"
COMPANY = ROOT / "content" / "company"
MONTH = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

SEAT_FILL = {"galaxium": ORANGE, "business": "1B1A2E", "economy": "8A88B5"}
# Where each body sits on the route map, from 0 to 1 across the drawing. Chosen so that no two routes cross.
# The last value puts the label above the node instead of below it.
MAP_AT = {"Venus": (0.03, 0.55, False), "Earth": (0.22, 0.55, False), "Moon": (0.30, 0.14, True), "Mars": (0.50, 0.26, True), "Jupiter": (0.76, 0.26, True),
          "Europa": (0.88, 0.60, False), "Pluto": (0.97, 0.88, False)}


def load(name: str):
    return yaml.safe_load((COMPANY / f"{name}.yaml").read_text(encoding="utf-8"))


def path(rel: str) -> Path:
    return DRIVE / rel


def long(d: date | str) -> str:
    d = date.fromisoformat(str(d)[:10])
    return f"{d.day} {MONTH[d.month]} {d.year}"


def ends(port: dict) -> str:
    """The end of a handling contract, as a reader says it."""
    value = str(port["contract_ends"])
    return long(value) if value[:1].isdigit() else {"rolling": "Rolling", "": "No contract: own staff"}[value]


class World:
    """The three files and the snapshot, with the questions the documents ask of them."""

    def __init__(self, d) -> None:
        self.fleet, self.ports, self.cabin = load("fleet"), load("ports"), load("cabin")
        self.classes = {c["id"]: c for c in self.fleet["classes"]}
        self.ships = self.fleet["ships"]
        self.crews = {c["name"]: c for c in self.fleet["crews"]}
        self.port = {p["body"]: p for p in self.ports}
        self.seat_classes = {c["id"]: c for c in self.cabin["classes"]}
        # One departure a month per route, always with the same block time: the first one stands for the route.
        self.routes = []
        for f in sorted(d.flights.values(), key=lambda f: f["flight_id"]):
            pair = (f["origin"], f["destination"])
            hours = round((datetime.fromisoformat(f["arrival_time"].replace("Z", "+00:00")) - datetime.fromisoformat(f["departure_time"].replace("Z", "+00:00"))).total_seconds() / 3600)
            known = next((r for r in self.routes if r["pair"] == pair), None)
            if known:
                assert known["hours"] == hours, f"block time of {pair} changes between departures"
                continue
            rotation = next((r for r in self.fleet["rotations"] if (r["origin"], r["destination"]) == pair), None)
            assert rotation, f"no ship and crew in fleet.yaml for {pair}"
            self.routes.append({"pair": pair, "origin": pair[0], "destination": pair[1], "hours": hours, "ship": rotation["ship"], "crew": rotation["crew"],
                                "base_price": f["base_price"], "seats": (f["economy_seats"], f["business_seats"], f["galaxium_seats"])})
        self.check()

    def check(self) -> None:
        seats = [s["class"] for row in self.fleet["cabin"]["rows"] for s in row["seats"]]
        assert len(seats) == self.fleet["cabin"]["seats"] == 10
        for r in self.routes:
            assert r["seats"] == tuple(seats.count(k) for k in ("economy", "business", "galaxium")), f"cabin plan and snapshot disagree on {r['pair']}"
            assert r["origin"] in self.port and r["destination"] in self.port, f"no port for {r['pair']}"
            assert r["crew"] in self.crews and any(s["name"] == r["ship"] for s in self.ships)
        assert len(self.routes) == len(self.fleet["rotations"]), "fleet.yaml has a rotation the snapshot does not fly"
        assert set(MAP_AT) == set(self.port)
        for s in self.ships:
            assert s["class"] in self.classes
        for c in self.cabin["classes"]:
            assert c["id"] in SEAT_FILL

    def called(self, body: str) -> str:
        """The body as the company says it: Luna, where the booking system says Moon."""
        return self.port[body]["called"]

    def route(self, r: dict) -> str:
        return f"{r['origin']} to {r['destination']}"

    def of_class(self, class_id: str) -> list[dict]:
        return [s for s in self.ships if s["class"] == class_id]

    def flown_by(self, names: list[str]) -> list[dict]:
        return [r for r in self.routes if r["ship"] in names]

    def at(self, body: str) -> list[tuple[str, dict]]:
        return [("Departs", r) for r in self.routes if r["origin"] == body] + [("Arrives", r) for r in self.routes if r["destination"] == body]

    def seat_rows(self) -> list[list]:
        return [[row["row"], ", ".join(s["id"] for s in row["seats"]), self.seat_classes[row["seats"][0]["class"]]["name"], row.get("note", "")] for row in self.fleet["cabin"]["rows"]]


SOURCE_SNAPSHOT = "Block hours: booking system. Ships and crews: roster of 20 October 2099"


# ------------------------------------------------------------------ the fleet

def fleet_deck(w: World) -> None:
    in_service = [s for s in w.ships if s["register"]["status"].startswith("In service")]
    deck = Deck(path("ops/fleet/fleet-guide.pptx"), "The fleet", "Three classes, six ships · Flight Operations · November 2099", owner="Priya Raman", modified="2099-11-02")
    deck.figures("Six ships, and every one of them has ten seats",
                 [(str(len(w.ships)), "ships"), (str(len(w.classes)), "classes"), (str(w.fleet["cabin"]["seats"]), "seats on each ship"), (str(len(in_service)), "in service today")],
                 items=["**The class decides where a ship may go**, not how many people she carries.", "**The cabin is the same on all six**, so a seat number means the same everywhere.",
                        "**What differs is behind and below the cabin:** shielding, coolant, life support, the hold and, on the Halcyon class, the rest deck."],
                 source="ops/fleet/fleet-register.xlsx")
    deck.table("The six ships", ["Ship", "Class", "In service since", "Home port", "Flies", "Status"],
               [[s["name"], w.classes[s["class"]]["name"], str(s["in_service_since"]), s["home_port"], s["flies"] or "–", s["register"]["status"]] for s in w.ships],
               widths=[1.2, 1.3, 1.1, 1.4, 3.6, 2.6], size=11, source="ops/fleet/fleet-register.xlsx")
    cs = list(w.classes.values())
    deck.table("The three classes side by side", [""] + [c["name"] for c in cs],
               [["Role"] + [c["role"] for c in cs], ["Certified to"] + [c["certified_to"] for c in cs], ["Length"] + [f"{c['length_m']} m" for c in cs],
                ["Crew"] + [f"{c['crew']}: {c['crew_roles'].lower()}" for c in cs], ["Hold"] + [f"{c['hold_kg']:,} kg" for c in cs],
                ["Life support"] + [f"{c['life_support_hours']} hours" for c in cs], ["Coolant loops"] + [str(c["coolant_loops"]) for c in cs],
                ["Observation"] + [c["observation"] for c in cs], ["Rest"] + [c["rest"] for c in cs]],
               widths=[1.2, 2.6, 2.6, 2.6], size=10.5, source="ops/fleet/data-sheets")
    kinds = {c["id"]: (c["name"], SEAT_FILL[c["id"]]) for c in reversed(w.cabin["classes"])}
    kinds["crew"] = ("Cabin lead", "")
    rows = []
    for row in w.fleet["cabin"]["rows"]:
        slots = [(s["id"], s["class"]) for s in row["seats"]]
        if row["row"] == 3:
            slots.append(("Cabin lead", "crew"))
        rows.append((f"Row {row['row']}", slots))
    deck.plan("The cabin: ten seats in six rows", rows, kinds, aft="Galley",
              aside=("The same on every ship", [w.fleet["cabin"]["overhead"], w.fleet["cabin"]["aft"]]), source="fleet data sheets, section 3")
    for c in cs:
        ships = w.of_class(c["id"])
        deck.bullets(f"{c['name']}: {c['role'] if c['role'].split()[0] in w.port else c['role'].lower()}", [c["summary"]] + [(f, 1) for f in c["fittings"]] + [f"**Limit.** {c['limits'][0]}"],
                     aside=("Ships", [f"**{s['name']}**, since {s['in_service_since']}. {s['note']}" for s in ships]), size=20, source=f"ops/fleet/data-sheets/{c['id']}-class.docx")
    deck.table("Who flies what on the scheduled service", ["Route", "Block hours", "Ship", "Crew", "Captain"],
               [[w.route(r), r["hours"], r["ship"], r["crew"], w.crews[r["crew"]]["captain"]] for r in w.routes], widths=[2.4, 1.1, 1.4, 1.2, 2.2], source=SOURCE_SNAPSHOT,
               note="The route names are the booking system's: it says Moon where we say Luna.")
    deck.timeline("How a leg is flown", [(p["phase"], p["what"].split(". ")[0].rstrip(".") + ".") for p in w.fleet["flight"]], items=w.fleet["flight_notes"])
    deck.bullets("Open items on the fleet", [item for c in cs for item in c["open_items"]], size=17,
                 source="ops/safety/bulletin-2099-07.md, ops/fleet/halcyon-maintenance-2099-10.md, ops/fleet/maintenance-plan-2099-q4.xlsx")
    deck.save()


def data_sheets(w: World) -> None:
    cabin = w.fleet["cabin"]
    for c in w.classes.values():
        ships = w.of_class(c["id"])
        doc = Doc(path(f"ops/fleet/data-sheets/{c['id']}-class.docx"), f"{c['name']}: data sheet", subtitle=c["role"], label="Operations · Fleet data sheet",
                  meta=[("Owner", "Zainab Hussein"), ("Updated", "30 October 2099"), ("Ships", ", ".join(s["name"] for s in ships)), ("Certified to", c["certified_to"])],
                  owner="Zainab Hussein", modified="2099-10-30")
        doc.h1("1. The class")
        doc.p(c["summary"])
        doc.fields([("Builder", c["builder"]), ("Length", f"{c['length_m']} m"), ("Seats", f"{cabin['seats']}: 6 economy, 3 business, 1 galaxium"), ("Crew", f"{c['crew']}: {c['crew_roles'].lower()}"),
                    ("Hold", f"{c['hold_kg']:,} kg"), ("Life support", f"{c['life_support_hours']} hours for travellers and crew"), ("Coolant loops", str(c["coolant_loops"])), ("Rest", c["rest"])])
        doc.h1("2. The ships")
        doc.table(["Ship", "In service since", "Home port", "Status", "Next check"],
                  [[s["name"], str(s["in_service_since"]), s["home_port"], s["register"]["status"], s["register"]["next_check"] or "–"] for s in ships], widths=[1.2, 1.2, 1.4, 2.6, 1.8])
        for s in ships:
            doc.p(f"**{s['name']}.** {s['note']}")
        doc.p("Status and next check are those of ops/fleet/fleet-register.xlsx.", size=9, italic=True)
        doc.h1("3. The cabin")
        doc.p("The cabin is the same on every ship of the fleet, so that a seat number means the same everywhere.")
        doc.table(["Row", "Seats", "Seat class", "Note"], w.seat_rows(), widths=[0.6, 1.2, 1.2, 3.6])
        doc.bullets([cabin["overhead"], cabin["aft"]])
        doc.h1("4. Fittings")
        doc.bullets(c["fittings"])
        n = 5
        if c.get("package"):
            doc.h1(f"{n}. The outer-system package")
            doc.p("What a hull of this class needs before it may fly beyond Jupiter:")
            doc.bullets(c["package"])
            n += 1
        doc.h1(f"{n}. Limits")
        doc.bullets(c["limits"])
        doc.h1(f"{n + 1}. Scheduled legs")
        legs = w.flown_by([s["name"] for s in ships])
        doc.table(["Route", "Block hours", "Ship", "Crew"], [[w.route(r), r["hours"], r["ship"], r["crew"]] for r in legs], widths=[2.6, 1.1, 1.4, 1.2])
        doc.p(f"The longest scheduled leg of the class is {max(r['hours'] for r in legs)} hours. {SOURCE_SNAPSHOT}.", size=9, italic=True)
        if c["open_items"]:
            doc.h1(f"{n + 2}. Open items")
            for item in c["open_items"]:
                doc.callout(item)
        doc.save()


def fleet_book(w: World) -> None:
    book = Book(path("ops/fleet/fleet-specifications.xlsx"), owner="Zainab Hussein", modified="2099-10-30")
    book.sheet("Ships", [{"ship": s["name"], "class": w.classes[s["class"]]["name"], "in_service_since": str(s["in_service_since"]), "home_port": s["home_port"], "flies": s["flies"],
                          "certified_to": s["register"]["certified_to"], "status": s["register"]["status"], "next_check": s["register"]["next_check"], "note": s["note"]} for s in w.ships],
               title="The fleet: ships", subtitle="Kept by Zainab Hussein. Status and next check agree with ops/fleet/fleet-register.xlsx.", status={"status": {"refit": "FBEBD3"}},
               widths={"flies": 46, "note": 70, "status": 44})
    book.sheet("Classes", [{"class": c["name"], "role": c["role"], "certified_to": c["certified_to"], "builder": c["builder"], "length_m": c["length_m"], "seats": w.fleet["cabin"]["seats"], "crew": c["crew"],
                            "crew_roles": c["crew_roles"], "hold_kg": c["hold_kg"], "life_support_hours": c["life_support_hours"], "coolant_loops": c["coolant_loops"],
                            "observation": c["observation"], "rest": c["rest"]} for c in w.classes.values()],
               title="The fleet: classes", subtitle="One row per class. The detail is in ops/fleet/data-sheets.", widths={"observation": 42, "rest": 42, "crew_roles": 46})
    book.sheet("Cabin plan", [{"row": row["row"], "seat": s["id"], "seat_class": w.seat_classes[s["class"]]["name"], "fare_factor": w.seat_classes[s["class"]]["factor"]}
                              for row in w.fleet["cabin"]["rows"] for s in row["seats"]],
               title="The cabin: ten seats, the same on every ship", subtitle="Fare factor: policies/pricing-policy.md. The fare is the base price of the flight times the factor.", formats={"fare_factor": "0.0"})
    book.sheet("Crews", [{"crew": c["name"], "size": c["size"], "captain": c["captain"], "ship": ", ".join(sorted({r["ship"] for r in w.routes if r["crew"] == c["name"]})),
                          "routes": ", ".join(w.route(r) for r in w.routes if r["crew"] == c["name"])} for c in w.crews.values()],
               title="Crews of the scheduled service", subtitle="Roster of 20 October 2099, kept by Andrei Popescu.", total={"size": "sum"}, widths={"routes": 58})
    book.sheet("Rotations", [{"route": w.route(r), "block_hours": r["hours"], "ship": r["ship"], "class": w.classes[next(s["class"] for s in w.ships if s["name"] == r["ship"])]["name"],
                              "crew": r["crew"]} for r in w.routes],
               title="Scheduled service: which ship flies which route", subtitle=SOURCE_SNAPSHOT + ".")
    book.notes("Notes", ["Seats, routes and block hours are those of the booking system. The route names are the booking system's: it says Moon where we say Luna.",
                         "Lengths, holds, life support and coolant loops are the builder's figures for the class.",
                         "Open items are in the data sheet of each class and in ops/fleet/maintenance-plan-2099-q4.xlsx."])
    book.close()


# ------------------------------------------------------------------ the ports

def legs_at(w: World, body: str) -> list[list]:
    return [[way, w.route(r), r["hours"], r["ship"], r["crew"]] for way, r in w.at(body)]


def port_briefs(w: World) -> None:
    for p in w.ports:
        same = p["called"] == p["body"]
        doc = Doc(path(f"ops/ground/port-briefs/port-brief-{p['slug']}.docx"), f"Port brief: {p['port']}", subtitle=f"{p['called']} · {p['kind']}", label="Operations · Port brief",
                  meta=[("Owner", "Lucia Bianchi"), ("Updated", "2 November 2099"), ("Handler", p["handler"]), ("Contract ends", ends(p))], owner="Lucia Bianchi", modified="2099-11-02")
        doc.h1("1. At a glance")
        doc.fields([("In the booking system", p["body"] if same else f"{p['body']} (we say {p['called']})"), ("Port", p["port"]), ("Gravity", p["gravity"]), ("Local time", p["local_time"]),
                    ("Our contact", p["contact"] or "The handler's duty desk"), ("Our presence", p["our_presence"])])
        doc.h1("2. Scheduled service here")
        doc.table(["", "Route", "Block hours", "Ship", "Crew"], legs_at(w, p["body"]), widths=[1, 2.4, 1.1, 1.3, 1.1])
        doc.p("One departure a month on each route. Contract departures, the Luna shuttle and positioning legs are not listed here: see ops/how-we-fly.md.", size=9, italic=True)
        doc.h1("3. The port")
        doc.bullets(p["facilities"])
        doc.h1("4. Beyond the port")
        doc.p("What the booking site describes to travellers, and how they get there from our ship.")
        doc.table(["Destination", "What it is"], [[b["name"], b["what"]] for b in p["beyond"]], widths=[1.6, 4.4])
        doc.p(f"**Transfer.** {p['transfer']}")
        doc.h1("5. What travellers need to know")
        doc.bullets(p["travellers"])
        doc.h1("6. What operations need to know")
        doc.bullets(p["operations"])
        if p["known"]:
            doc.h1("7. Known items")
            for item in p["known"]:
                doc.callout(item)
        doc.save()


def network_deck(w: World) -> None:
    deck = Deck(path("ops/ground/ports-and-destinations.pptx"), "Ports and destinations", "Where we dock, and where travellers are going · Ground Operations · November 2099",
                owner="Lucia Bianchi", modified="2099-11-02")
    deck.statement("Two names for every place", "The booking site sells a destination. Our ship docks at a port.",
                   "A traveller books Mars and reads about Olympus Base. Operations says Olympus Port, because that is where the pad, the handler and the contract are. Both names are right, and this deck puts them side by side.")
    nodes = {body: (x, y, w.called(body), w.port[body]["port"], above) for body, (x, y, above) in MAP_AT.items()}
    links, seen = [], set()
    for r in sorted(w.routes, key=lambda r: r["origin"] != "Earth"):  # "out" is the leg that leaves Earth
        if r["pair"] in seen:
            continue
        back = next((x for x in w.routes if x["pair"] == r["pair"][::-1]), None)
        seen.update({r["pair"], r["pair"][::-1]})
        label = f"{r['hours']} h" if not back or back["hours"] == r["hours"] else f"{r['hours']} h out, {back['hours']} h back"
        links.append((r["origin"], r["destination"], label, bool(back)))
    deck.network("The scheduled network: ten routes between seven ports", nodes, links, source="Booking system (routes and block hours)",
                 note="A heavy line is flown in both directions, an arrow in one. Not to scale: every route has a fixed block time, whatever the distance on the day.")
    deck.table("Where we dock, and what lies beyond", ["Booking system", "We say", "Port", "Beyond the port"],
               [[p["body"], p["called"], p["port"], ", ".join(b["name"] for b in p["beyond"])] for p in w.ports], widths=[1.3, 1, 2.2, 3.6], source="ops/ground/port-briefs")
    deck.table("Who handles us, and until when", ["Port", "Handler", "Contract ends", "Our presence"],
               [[f"{p['called']}, {p['port']}", p["handler"], ends(p), p["our_presence"]] for p in w.ports], widths=[2.2, 2.4, 1.6, 4.6], size=10.5, source="ops/ground/station-contacts.md")
    deck.table("What it is like there", ["Port", "What kind of place", "Gravity", "Local time"],
               [[p["port"], p["kind"], p["gravity"], p["local_time"]] for p in w.ports], widths=[1.7, 3.6, 1.8, 3.2], size=10.5, source="ops/ground/port-briefs")
    for p in w.ports:
        beyond = [f"**{b['name']}.** {b['what']}" for b in p["beyond"]] + [f"**Transfer.** {p['transfer']}"]
        deck.two(f"{p['called']}: {p['port']}", ("The port", [p["kind"] + "."] + p["facilities"]), ("Beyond it", beyond + [(t, 1) for t in p["travellers"][:2]]), size=16,
                 source=f"ops/ground/port-briefs/port-brief-{p['slug']}.docx")
    deck.save()


def port_directory(w: World) -> None:
    book = Book(path("ops/ground/port-directory.xlsx"), owner="Lucia Bianchi", modified="2099-11-02")
    book.sheet("Ports", [{"booking_system": p["body"], "we_say": p["called"], "port": p["port"], "kind": p["kind"], "handler": p["handler"], "contract_ends": str(p["contract_ends"]),
                          "our_contact": p["contact"], "gravity": p["gravity"], "local_time": p["local_time"], "departures_a_month": sum(r["origin"] == p["body"] for r in w.routes),
                          "arrivals_a_month": sum(r["destination"] == p["body"] for r in w.routes), "our_presence": p["our_presence"]} for p in w.ports],
               title="Ports of the scheduled service", subtitle="Kept by Lucia Bianchi. Handlers and contract dates: ops/ground/station-contacts.md. Departures and arrivals: scheduled service only, from the booking system.",
               total={"departures_a_month": "sum", "arrivals_a_month": "sum"}, widths={"kind": 52, "our_presence": 70, "local_time": 44, "handler": 36})
    book.sheet("Destinations", [{"booking_system": p["body"], "port": p["port"], "destination": b["name"], "what_it_is": b["what"], "transfer": p["transfer"]} for p in w.ports for b in p["beyond"]],
               title="What lies beyond each port", subtitle="The destination names are those of the public booking site.", widths={"what_it_is": 70, "transfer": 70})
    book.sheet("Routes", [{"route": w.route(r), "from_port": w.port[r["origin"]]["port"], "to_port": w.port[r["destination"]]["port"], "block_hours": r["hours"], "ship": r["ship"], "crew": r["crew"],
                           "handler_at_origin": w.port[r["origin"]]["handler"], "handler_at_destination": w.port[r["destination"]]["handler"]} for r in w.routes],
               title="Scheduled routes, port to port", subtitle=SOURCE_SNAPSHOT + ".", widths={"handler_at_origin": 34, "handler_at_destination": 34})
    book.notes("Notes", ["The booking system names the body (Moon, Mars). Operations names the port (Tycho Spaceport, Olympus Port). The public site describes the destination beyond the port (Artemis Base Camp, Olympus Base).",
                         "Transfers beyond the port are run by the port or the station and are not part of our ticket.", "Turnaround counts and service levels by port are in ops/ground/turnaround-scorecard-2099-q3.xlsx."])
    book.close()


# ------------------------------------------------------------------ the cabin

def cabin_guide(w: World) -> None:
    every = w.cabin["for_everyone"]
    doc = Doc(path("marketing/product/cabin-class-guide.docx"), "Cabin class guide", subtitle="Economy, business and galaxium: what we promise and what is behind it", label="Marketing · Product guide",
              meta=[("Owner", "Noor Rahimi"), ("Checked by", "Priya Raman, Nadia Haddad"), ("Updated", "8 October 2099"), ("For", "Sales, Support and anyone writing about the product")],
              owner="Noor Rahimi", modified="2099-10-08")
    doc.callout("The words in bold under each class are the booking site's. Use them as they are. If a campaign needs other words, the site changes first.", "House rule")
    doc.h1("1. The three classes")
    doc.table(["Seat class", "Seats", "On every ship", "Fare"], [[c["name"], c["seats"], sum(s["class"] == c["id"] for row in w.fleet["cabin"]["rows"] for s in row["seats"]),
                                                                  "The base price" if c["factor"] == 1 else f"{c['factor']} times the base price"] for c in w.cabin["classes"]], widths=[1.3, 1.6, 1.1, 2.4])
    doc.p("Ten seats on every ship and every departure. The fare is the base price of the flight times the factor of the class: policies/pricing-policy.md.", size=9, italic=True)
    for k, c in enumerate(w.cabin["classes"], 2):
        doc.h1(f"{k}. {c['name']}")
        doc.p("**" + " · ".join(c["promise"]) + "**")
        doc.bullets(c["detail"])
    doc.h1("5. For every traveller, in every class")
    doc.table(["", ""], [["Baggage", every["baggage"]], ["Meals", every["meals"]], ["Medical equipment", every["medical"]], ["Assistance", every["assistance"]], ["Seat holds", every["holds"]], ["The way home", every["way_home"]], ["The Luna shuttle", every["connection"]]], widths=[1.4, 4.8])
    doc.h1("6. The cabin")
    doc.table(["Row", "Seats", "Seat class", "Note"], w.seat_rows(), widths=[0.6, 1.2, 1.2, 3.6])
    doc.p("The plan is drawn in ops/fleet/fleet-guide.pptx. The ships differ behind and below the cabin, not in it: see the data sheets in ops/fleet/data-sheets.")
    doc.h1("7. What we do not say")
    doc.bullets(["That a flight is weightless. The cabin has weight for most of the flight; it is weightless at turnover, for about twenty minutes.",
                 "That there is a lounge everywhere. There is one at Cape Meridian and one at Tycho.", "That transfers beyond the port are ours. They are run by the port or the station."])
    doc.save()


def traveller_guide(w: World) -> None:
    every = w.cabin["for_everyone"]
    doc = Doc(path("support/knowledge-base/before-you-fly.docx"), "Before you fly", subtitle="What we tell travellers, from booking to arrival", label="Support · Knowledge base",
              meta=[("Owner", "Greta Olsen"), ("Updated", "22 October 2099"), ("Use", "Answers to travellers. Copy the wording."), ("Questions", "#support-escalations")], owner="Greta Olsen", modified="2099-10-22")
    doc.h1("1. From booking to arrival")
    doc.table(["When", "What we tell the traveller"], [[s["when"], s["what"]] for s in w.cabin["before_you_fly"]], widths=[1.4, 4.8])
    doc.h1("2. The same for every traveller")
    doc.table(["", ""], [["Baggage", every["baggage"]], ["Meals", every["meals"]], ["Medical equipment", every["medical"]], ["Assistance", every["assistance"]], ["Seat holds", every["holds"]], ["The way home", every["way_home"]], ["The Luna shuttle", every["connection"]]], widths=[1.4, 4.8])
    doc.callout("On the Kestrel class ships heavy items go under the seat, not in the overhead bins, until the restraint latches are replaced (safety bulletin 2099-07).", "Until further notice")
    doc.h1("3. On board")
    doc.table(["Part of the flight", "How long", "What happens"], [[p["phase"], p["about"], p["what"]] for p in w.fleet["flight"]], widths=[1.5, 1.3, 3.6])
    doc.bullets(w.fleet["flight_notes"])
    doc.h1("4. By destination")
    doc.p("The booking site names the destination, our crews name the port. A traveller may use either.")
    for p in w.ports:
        doc.h2(f"{p['body']}: {', '.join(b['name'] for b in p['beyond'])}" if p["beyond"][0]["name"] != p["body"] else p["body"])
        doc.p(f"We dock at **{p['port']}**. {p['transfer']}")
        doc.bullets(p["travellers"])
    doc.h1("5. Who to ask")
    doc.bullets(["Assistance, baggage, anything at a port: Operations, in #support-escalations, naming the port.", "What a seat class includes: marketing/product/cabin-class-guide.docx.",
                 "What a port is like: the port briefs in ops/ground/port-briefs."])
    doc.save()


def build(d) -> None:
    w = World(d)
    fleet_deck(w)
    data_sheets(w)
    fleet_book(w)
    port_briefs(w)
    network_deck(w)
    port_directory(w)
    cabin_guide(w)
    traveller_guide(w)
