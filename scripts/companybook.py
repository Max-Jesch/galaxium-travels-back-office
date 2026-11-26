"""The company as a business: who owns it, how it earns, what it costs and how a year of
flying is made up.

`documents.build()` calls `build()`. Two files land in `content/generated/drive/`:

    leadership/company-profile.docx                      the company on six pages
    finance/annual/financial-summary-2094-2099.xlsx      revenue, costs and result by year

It also writes `docs/company.html`: the same facts on one page for people working on the
repo. The page is plain HTML with no script, so it opens in any viewer.

Every fact is in `content/company/finance.yaml`; nothing about money or ownership is
typed in here. Fares, routes and seats are read from the snapshot. The rules of
`documents.py` apply. `check()` fails the build if a year does not add up or if the file
stops agreeing with the snapshot.
"""

from __future__ import annotations

from collections import Counter
from html import escape as e
from pathlib import Path

import yaml

from office import Book, Doc

ROOT = Path(__file__).resolve().parent.parent
DRIVE = ROOT / "content" / "generated" / "drive"
FACTORS = {"economy": 1, "business": 2.5, "galaxium": 5}
PLACE = {"Moon": "Luna", "Earth": "Earth", "Mars": "Mars", "Jupiter": "Jupiter", "Pluto": "Pluto"}  # the pattern's bodies, in the words of way_home
GROUPS = {"Flying": ["Operations"], "Commercial": ["Sales", "Marketing", "Support"], "Digital": ["Engineering", "IT"], "Corporate": ["Leadership", "Finance", "Legal", "People"]}

CSS = """
  :root {
    color-scheme: light dark;
    --fg: var(--foreground, CanvasText);
    --dim: var(--muted-foreground, color-mix(in srgb, CanvasText 60%, transparent));
    --line: var(--border, color-mix(in srgb, CanvasText 18%, transparent));
    --box: var(--card, color-mix(in srgb, CanvasText 4%, transparent));
    --bar: var(--chart-1, #4f6bd8);
    --track: color-mix(in srgb, var(--fg) 8%, transparent);
    --r: var(--radius, 8px);
  }
  body { margin: 0; font: 13px/1.45 var(--font-sans, system-ui, sans-serif); color: var(--fg); }
  .wrap { max-width: 1000px; margin: 0 auto; }
  .premise { font-size: 15px; margin: 0 0 6px; }
  h2 { font-size: 14px; margin: 28px 0 4px; }
  .note { margin: 0 0 10px; color: var(--dim); }
  .cols { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 18px 28px; }
  .cap { font-size: 11px; font-weight: 600; color: var(--dim); text-transform: uppercase; letter-spacing: .04em; margin: 0 0 6px; }
  .row { display: grid; grid-template-columns: minmax(120px, 44%) 1fr 64px; gap: 8px; align-items: center; padding: 3px 0; }
  .row .name { line-height: 1.25; }
  .row .num { text-align: right; font-variant-numeric: tabular-nums; }
  .track { height: 10px; background: var(--track); border-radius: 4px; }
  .fill { height: 10px; background: var(--bar); border-radius: 4px; min-width: 2px; }
  .sum { border-top: 1px solid var(--line); margin-top: 4px; padding-top: 5px; font-weight: 600; }
  table { border-collapse: collapse; width: 100%; font-variant-numeric: tabular-nums; }
  th { text-align: right; font-size: 11px; font-weight: 600; color: var(--dim); padding: 4px 6px; border-bottom: 1px solid var(--line); }
  td { text-align: right; padding: 5px 6px; border-bottom: 1px solid var(--line); }
  th:first-child, td:first-child, td.t, th.t { text-align: left; }
  td.t { color: var(--dim); font-size: 12px; }
  .scroll { overflow-x: auto; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 8px; }
  .card { border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; }
  .card b { display: block; }
  .card .who { color: var(--dim); font-size: 11.5px; display: block; }
  .card p { margin: 5px 0 0; font-size: 12px; }
  .card span.i { display: block; font-size: 12px; }
  .own { display: flex; height: 12px; gap: 2px; margin: 4px 0 8px; }
  .own div { background: var(--bar); border-radius: 3px; }
  .own div:nth-child(2) { opacity: .7; } .own div:nth-child(3) { opacity: .45; }
  .tl { display: grid; grid-template-columns: 48px 1fr; gap: 4px 10px; }
  .tl b { font-variant-numeric: tabular-nums; }
"""


def path(rel: str) -> Path:
    return DRIVE / rel


def load() -> dict:
    return yaml.safe_load((ROOT / "content" / "company" / "finance.yaml").read_text(encoding="utf-8"))


def label(y: dict) -> str:
    return f"{y['year']} forecast" if y.get("forecast") else str(y["year"])


def routes(d) -> list[dict]:
    """The scheduled routes of the snapshot, with what one full departure earns."""
    seen = {}
    for f in sorted(d.flights.values(), key=lambda f: f["flight_id"]):
        pair = (f["origin"], f["destination"])
        if pair not in seen:
            full = sum(int(f[f"{k}_seats"]) * int(f[f"{k}_price"]) for k in FACTORS)
            seen[pair] = {"route": f"{pair[0]} to {pair[1]}", "economy": int(f["economy_price"]), "full": full,
                          "whole_ship": int(int(f["economy_price"]) * 2.5 * 10 * 0.8)}  # the charter rule of the rate card
    return list(seen.values())


def check(f: dict, d) -> None:
    assert sum(o["share"] for o in f["owners"]) == 100, "the owners' shares do not add up to 100"
    for y in f["years"]:
        assert len(y["revenue"]) == len(f["revenue_lines"]) and len(y["costs"]) == len(f["cost_lines"]), f"{y['year']}: a line is missing"
        assert y["seats_sold"] <= y["seats_offered"]
    fly = {x["year"]: x for x in f["flying"]}
    assert [y["year"] for y in f["years"]] == list(fly), "finance.yaml: `years` and `flying` cover different years"
    for y in f["years"]:
        assert y["seats_offered"] == fly[y["year"]]["scheduled"] * 10, f"{y['year']}: seats offered is not ten a scheduled leg"
    assert f["flying"][-1]["routes"] == len(routes(d)), "finance.yaml and the snapshot disagree on the number of routes"
    assert f["flying"][-1]["ships"] == 6
    # The flying programme has to agree with the monthly pattern of fleet.yaml.
    pattern = yaml.safe_load((ROOT / "content" / "company" / "fleet.yaml").read_text(encoding="utf-8"))["pattern"]
    kinds = Counter(x["kind"] for x in pattern["legs"])
    last = f["flying"][-1]
    assert last["shuttle"] == 12 * (2 * len(pattern["shuttle"]["mornings"]) + kinds["Luna shuttle"]), "finance.yaml: shuttle legs do not follow from fleet.yaml's pattern"
    assert last["positioning"] == 12 * kinds["Positioning"] - 3, "finance.yaml: positioning legs do not follow from the pattern (Pluto flew nine months of 2099)"
    assert last["charter"] >= 12 * (kinds["Contract"] + 2 * len(pattern["contract_mornings"])), "finance.yaml: fewer contract legs than the pattern flies"
    for way in f["way_home"]["legs"]:
        assert any(x["kind"] != "Contract" and PLACE.get(x["origin"]) == way["from"] and PLACE.get(x["destination"]) == way["to"] for x in pattern["legs"]), f"no leg in the pattern takes travellers home from {way['from']}"


def build(d) -> None:
    f = load()
    check(f, d)
    years, rs = f["years"], routes(d)
    total = lambda y: sum(y["revenue"])  # noqa: E731
    result = lambda y: sum(y["revenue"]) - sum(y["costs"])  # noqa: E731
    fly = {x["year"]: x for x in f["flying"]}
    last = next(y for y in reversed(years) if not y.get("forecast"))  # the last closed year
    mars, pluto = next(r for r in rs if r["route"] == "Earth to Mars"), next(r for r in rs if r["route"] == "Earth to Pluto")
    first_profit = next(y for y in years if result(y) > 0)

    # ---- the profile
    doc = Doc(path("leadership/company-profile.docx"), "Galaxium Travels: company profile", subtitle="Who we are, how we earn and what we depend on", label="Internal",
              meta=[("Owner", "Renata Castellanos"), ("Figures", "Finance, at the August close"), ("Amounts", "US dollars")], owner="Renata Castellanos", modified="2099-09-07")
    doc.h1("1. What we are")
    doc.p(f"Galaxium Travels is a scheduled carrier. We fly {len(rs)} routes between Earth and Pluto with six ships of ten seats, one departure a month on each route, "
          "and we fly charter rotations for companies that work out there. We fly and maintain the ships. We do not build them, and we do not own the ports.")
    doc.p(f["position"]["what"])
    doc.h1("2. Why we were founded")
    doc.p(f["why_founded"])
    doc.h1("3. How we earn")
    doc.table(["Line", "Who buys", "What is sold", f"Revenue {last['year']}, $ million"],
              [["Scheduled service", "Travellers, on the public site", "One seat on one departure", last["revenue"][0]],
               ["Charter rotations", "Companies, under a contract", "A block of seats or a whole ship: on a contract departure, on the Luna shuttle or on a leg we fly anyway", last["revenue"][1]],
               ["Cargo and other", "Forwarders, travellers and companies", "Freight in the hold of the Luna shuttle, hold space booked ahead, excess baggage", last["revenue"][2]],
               ["Total", "", "", total(last)]], widths=[1.6, 2.2, 3.4, 1.6], total=True)
    doc.bullets([f"A seat costs the route's base price times the seat class factor: economy 1, business 2.5, galaxium 5. A full departure earns 18.5 times the economy fare.",
                 f"Earth to Mars: ${mars['economy']:,} in economy, ${mars['full']:,} for a full ship. Earth to Pluto: ${pluto['economy']:,} and ${pluto['full']:,}.",
                 f"A whole ship on charter is priced as ten business seats less 20%: ${mars['whole_ship']:,} from Earth to Mars. Few contract legs are whole ships. Most carry a block of a few seats, "
                 f"and a contract leg earned ${1000 * last['revenue'][1] // fly[last['year']]['charter']:,},000 on average in {last['year']}.",
                 f"**The way home.** {f['way_home']['rule']} {f['way_home']['how']}",
                 "The Luna shuttle flies ten nights a month. Its hold is sold as freight and its seats under contract. It is not on the public site: the public seats between Earth and Luna belong to the Free Port's franchise carriers, and we hold one public slot a month."])
    doc.h1("4. What we depend on")
    doc.table(["What", "Whose it is", "How we pay"], [[x["what"], x["who"], x["how"]] for x in f["depends_on"]], widths=[1.6, 1.8, 5.2])
    doc.h1(f"5. What it costs ({last['year']})")
    doc.table(["Cost", "$ million", "Share of revenue"], [[line, c, f"{c / total(last):.0%}"] for line, c in zip(f["cost_lines"], last["costs"])]
              + [["Total", sum(last["costs"]), f"{sum(last['costs']) / total(last):.0%}"]], widths=[4, 1.4, 1.6], total=True)
    doc.p(f"The ships are the largest cost and they are fixed: a lease runs whether a ship flies or not. That is why a ship in refit costs twice, and why every seat left empty is lost for good.")
    doc.h1("6. The years")
    doc.table(["Year", "Routes", "Ships", "Seats offered", "Seats sold", "Revenue, $ million", "Result, $ million"],
              [[label(y), fly[y["year"]]["routes"], fly[y["year"]]["ships"], y["seats_offered"], y["seats_sold"], total(y), result(y)] for y in years], widths=[1.5, 0.9, 0.8, 1.2, 1.1, 1.5, 1.5])
    doc.bullets([f"**{y['year']}.** {y['note']}" for y in years])
    doc.callout(f"{first_profit['year']} was the first year with a profit. Until then the company was paid for by its owners.", "Where we stand")
    doc.h1("7. Who owns us")
    doc.table(["Owner", "Share", ""], [[o["name"], f"{o['share']}%", o["note"]] for o in f["owners"]], widths=[2.2, 0.8, 5.6])
    doc.table(["Year", "Raised, $ million", "For what"], [[str(x["year"]), x["raised"], x["what"]] for x in f["funding"]] + [["Total", sum(x["raised"] for x in f["funding"]), ""]],
              widths=[0.9, 1.5, 6.2], total=True)
    doc.p(f"**The board.** {f['board']['seats']} {f['board']['meets']}")
    doc.table(["Seat", "Who", "From"], [[m["seat"], m["name"], m["from"]] for m in f["board"]["members"]], widths=[1.2, 2.4, 5])
    doc.table(["Company", "Where", "What it does"], [[e["name"], e["where"], e["what"]] for e in f["entities"]], widths=[2.4, 2, 4.2])
    doc.h1("8. Who else flies")
    doc.p(f["position"]["market"])
    doc.table(["Carrier", "What they are", "What we do that they do not"], [[c["name"], c["they"], c["we"]] for c in f["position"]["against"]], widths=[1.6, 3.6, 3.4])
    doc.h1("9. How a year of flying is made up")
    doc.table(["Kind", "Slot", "What it is", "How it is sold"], [[k["kind"], k["slot"], k["what"], k["sold"]] for k in f["flying_kinds"]], widths=[1.5, 1.2, 3.3, 2.8])
    doc.table(["Year", "Scheduled legs", "Contract legs", "Luna shuttle legs", "Positioning legs", "All legs"],
              [[label(y), x["scheduled"], x["charter"], x["shuttle"], x["positioning"], x["scheduled"] + x["charter"] + x["shuttle"] + x["positioning"]] for y, x in ((y, fly[y["year"]]) for y in years)],
              widths=[1.4, 1.3, 1.3, 1.4, 1.4, 1.2])
    doc.p(f"{f['slots']['rule']} {f['slots']['why_it_matters']}")
    doc.table(["Slot", "Who may be on board", "How we get one"], [[k["kind"], k["carries"], k["how"]] for k in f["slots"]["kinds"]], widths=[1.3, 3, 4.5])
    doc.p(f"**Why one a month.** {f['slots']['why_one_a_month']} {f['slots']['next']}")
    doc.p(f"**Sales books.** {f['books']['rule']} {f['books']['follows']}")
    doc.save()

    # ---- the workbook
    book = Book(path("finance/annual/financial-summary-2094-2099.xlsx"), owner="Ingrid Solberg", modified="2099-09-18", title="Financial summary 2094 to 2099")
    key = lambda s: s.lower().replace(",", "").replace(" ", "_").replace("-", "_")  # noqa: E731
    book.sheet("Summary", [{"year": label(y), "revenue": total(y), "costs": sum(y["costs"]), "result": result(y), "result_share_of_revenue": round(result(y) / total(y), 3), "note": y["note"]} for y in years],
               title="Revenue, costs and result by year", subtitle="Millions of dollars. Calendar years. 2099 is the forecast made at the August close.",
               formats={"result_share_of_revenue": "0.0%"}, widths={"note": 90})
    book.sheet("Revenue", [{"year": label(y), **{key(line): v for line, v in zip(f["revenue_lines"], y["revenue"])}, "total": total(y)} for y in years],
               title="Revenue by line", subtitle="Millions of dollars. Scheduled service from the sales books, charter from the contracts, by the year flown.")
    book.sheet("Costs", [{"year": label(y), **{key(line): v for line, v in zip(f["cost_lines"], y["costs"])}, "total": sum(y["costs"])} for y in years],
               title="Costs by line", subtitle="Millions of dollars.")
    book.sheet("Flying", [{"year": label(y), "routes": x["routes"], "ships": x["ships"], "scheduled_legs": x["scheduled"], "contract_legs": x["charter"], "luna_shuttle_legs": x["shuttle"], "positioning_legs": x["positioning"],
                           "seats_offered": y["seats_offered"], "seats_sold": y["seats_sold"], "share_sold": round(y["seats_sold"] / y["seats_offered"], 3)} for y, x in ((y, fly[y["year"]]) for y in years)],
               title="Flying and seats by year", subtitle="A leg is one flight from port to port. Seats are those of the scheduled service: ten a leg.", formats={"share_sold": "0%"})
    book.sheet("Owners and funding", [{"item": o["name"], "share_or_amount": f"{o['share']}%", "note": o["note"]} for o in f["owners"]]
               + [{"item": f"Round of {x['year']}", "share_or_amount": f"${x['raised']} million", "note": x["what"]} for x in f["funding"]],
               title="Who owns the company, and what was raised", widths={"note": 90})
    book.notes("Notes", ["Prepared by Finance for the board and the auditors. Closed years are from the ledger.", "2099 is the forecast made at the August close. It is replaced when the year is closed.",
                         f["books"]["rule"], f["books"]["follows"], "Charter flying is sold under contract and is not in the booking system."])
    book.close()
    page(f, d, rs)


def page(f: dict, d, rs: list[dict]) -> None:
    """docs/company.html: the company on one page, written out as plain HTML."""
    company = yaml.safe_load((ROOT / "content" / "company" / "company.yaml").read_text(encoding="utf-8"))
    crews = yaml.safe_load((ROOT / "content" / "company" / "fleet.yaml").read_text(encoding="utf-8"))["crews"]
    years, fly = f["years"], {x["year"]: x for x in f["flying"]}
    closed = next(y for y in reversed(years) if not y.get("forecast"))
    rev, cost = sum(closed["revenue"]), sum(closed["costs"])
    out: list[str] = []
    add = out.append

    def sec(title: str, note: str = "") -> None:
        add(f"<h2>{e(title)}</h2>" + (f'<p class="note">{e(note)}</p>' if note else ""))

    def bars(cap: str, rows: list[tuple[str, int]], scale: int, total: str | None = None) -> str:
        html = f'<div><p class="cap">{e(cap)}</p>'
        for name, value in rows:
            html += (f'<div class="row"><span class="name">{e(name)}</span><div class="track"><div class="fill" style="width:{100 * value / scale:.1f}%"></div></div>'
                     f'<span class="num">{value:,}</span></div>')
        if total:
            html += f'<div class="row sum"><span class="name">{e(total)}</span><span></span><span class="num">{sum(v for _, v in rows):,}</span></div>'
        return html + "</div>"

    def cards(items: list[tuple[str, str, str]]) -> None:
        add('<div class="grid">' + "".join(f'<div class="card"><b>{e(a)}</b><span class="who">{e(b)}</span>' + (f"<p>{e(c)}</p>" if c else "") + "</div>" for a, b, c in items) + "</div>")

    add(f'<p class="premise">A scheduled carrier with six leased ships of ten seats. It sells seats to travellers on {len(rs)} routes from Earth to Pluto and flies charter rotations '
        "for companies. It flies and maintains the ships; it does not build them and does not own the ports.</p>")
    add(f'<p class="note">{e(f["why_founded"])}</p>')

    sec(f"Where a dollar comes from and where it goes ({closed['year']}, $ million)",
        f"Revenue {rev:,}, costs {cost:,}, result {rev - cost:,}: {(rev - cost) / rev:.1%} of revenue.")
    scale = max(closed["revenue"] + closed["costs"])
    add('<div class="cols">' + bars("Who pays us", list(zip(f["revenue_lines"], closed["revenue"])), scale, "Revenue")
        + bars("Whom we pay", list(zip(f["cost_lines"], closed["costs"])), scale, "Costs") + "</div>")

    sec("What one sale is worth", "A seat costs the route's base price times the class factor: economy 1, business 2.5, galaxium 5. A whole ship on charter is ten business seats less 20%.")
    r = {x["route"]: x for x in rs}
    cards([(f"${r['Earth to Moon']['economy']:,}", "Earth to Moon, economy seat", ""), (f"${r['Earth to Mars']['full']:,}", "Earth to Mars, full departure", ""),
           (f"${r['Earth to Mars']['whole_ship']:,}", "Earth to Mars, whole ship on charter", ""), (f"${r['Earth to Pluto']['full']:,}", "Earth to Pluto, full departure", "")])

    sec("What it depends on", "Almost everything except the flying, the maintenance and the selling belongs to somebody else.")
    cards([(x["what"], x["who"], x["how"]) for x in f["depends_on"]])

    sec("The years ($ million)", "The last year is the forecast made at the August close.")
    add('<div class="scroll"><table><tr><th>Year</th><th>Routes</th><th>Ships</th><th>Seats sold</th><th>Of offered</th><th>Revenue</th><th>Result</th><th class="t"></th></tr>')
    for y in years:
        res = sum(y["revenue"]) - sum(y["costs"])
        add(f"<tr><td>{label(y)}</td><td>{fly[y['year']]['routes']}</td><td>{fly[y['year']]['ships']}</td><td>{y['seats_sold']:,}</td><td>{y['seats_sold'] / y['seats_offered']:.0%}</td>"
            f"<td>{sum(y['revenue']):,}</td><td>{res:+,}</td><td class=\"t\">{e(y['note'])}</td></tr>")
    add("</table></div>")

    sec("How a year of flying is made up", f"{f['slots']['rule']} {f['slots']['why_it_matters']}")
    last = f["flying"][-1]
    add('<div class="cols">' + bars(f"Legs flown, {label(years[-1])}", [("Scheduled, sold on the site", last["scheduled"]), ("Contract, sold by Sales", last["charter"]),
                                                                            ("Luna shuttle: freight, contract seats", last["shuttle"]), ("Positioning, not sold", last["positioning"])], last["shuttle"], "All legs")
        + f'<div><p class="cap">Why the booking system starts on 10 August</p><p class="note">{e(f["books"]["rule"])} {e(f["books"]["follows"])}</p></div></div>')

    add('<div class="scroll"><table><tr><th>Slot</th><th class="t">Who may be on board</th><th class="t">How we get one</th></tr>'
        + "".join(f'<tr><td>{e(k["kind"])}</td><td class="t">{e(k["carries"])}</td><td class="t">{e(k["how"])}</td></tr>' for k in f["slots"]["kinds"]) + "</table></div>")
    add(f'<p class="note" style="margin-top:8px">{e(f["slots"]["why_one_a_month"])} {e(f["slots"]["next"])}</p>')

    way = f["way_home"]
    sec("How a traveller gets home", f"{way['rule']} {way['how']}")
    cards([(f"{x['from']} to {x['to']}", "Included in the fare", x["ride"]) for x in way["legs"]]
          + [("Mars, Venus, Europa", "A scheduled departure", way["from_mars_venus_europa"]), ("The Luna shuttle", "Free with a ticket", way["luna_connection"]),
             ("Under contract", "Half the outbound rate", way["under_contract"])])
    sec("Where we fly under contract only", "Not on the public site. The rate card of Sales has the prices.")
    cards([(x["name"], x["where"], x["what"]) for x in f["contract_ports"]])

    raised = ", ".join(f"{x['year']} ${x['raised']}m" for x in f["funding"])
    sec("Who owns it", f"{f['board']['seats']} {f['board']['meets']} Raised ${sum(x['raised'] for x in f['funding']):,} million in three rounds: {raised}.")
    add('<div class="own">' + "".join(f'<div style="width:{o["share"]}%"></div>' for o in f["owners"]) + "</div>")
    cards([(f"{o['share']}%", o["name"], o["note"]) for o in f["owners"]])
    add('<p class="cap" style="margin-top:12px">The board</p>')
    cards([(m["name"], m["seat"], m["from"]) for m in f["board"]["members"]])

    dept, site = Counter(p["department"] for p in d.people), Counter(p["location"] for p in d.people)
    crew = sum(c["size"] for c in crews)
    where = ", ".join(f"{k} ({v})" for k, v in site.most_common())
    heads, counted = [p for p in d.everyone if p.get("kind") == "station"], sum(x["people"] for x in company["workforce"])
    assert sum(1 for p in d.everyone if p.get("kind") == "crew") == crew, "people.yaml and fleet.yaml disagree on the flight crew"
    sec("How it is organised", f"{len(d.people) + crew + len(heads) + counted} people. {len(d.people)} office staff: {where}. {crew} flight crew in {len(crews)} crews. "
        f"{len(heads) + counted} at the hangar, on the ramp and at the station desks. {company['standby']} Two companies: " + " and ".join(f"{x['name']} ({x['where']})" for x in f["entities"]) + ".")
    add('<div class="grid">' + "".join(f'<div class="card"><span class="who">{e(g)}</span>' + "".join(f'<span class="i">{e(k)} {dept[k]}</span>' for k in names)
                                       + (f'<span class="i">Flight crew {crew}</span>' if g == "Flying" else "") + "</div>" for g, names in GROUPS.items()) + "</div>")
    assert sum(dept[k] for names in GROUPS.values() for k in names) == len(d.people), "a department is missing from GROUPS in companybook.py"
    add('<div class="scroll" style="margin-top:10px"><table><tr><th>Hangar and stations</th><th>People</th><th class="t">What they do</th><th class="t">Led by</th><th class="t">Paid from</th></tr>'
        + "".join(f'<tr><td>{e(x["where"])}</td><td>{x["people"]}</td><td class="t">{e(x["what"])}</td><td class="t">{e(x["lead"])}</td><td class="t">{e(x["paid_from"])}</td></tr>' for x in company["workforce"])
        + "</table></div>")

    sec("Where it stands", f["position"]["what"] + " " + f["position"]["market"])
    cards([(x["name"], x["they"], "We: " + x["we"]) for x in f["position"]["against"]])

    sec("History")
    add('<div class="tl">' + "".join(f"<b>{e(str(h['when']))}</b><span>{e(h['what'])}</span>" for h in company["history"]) + "</div>")

    html = ("<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
            "<title>Galaxium Travels: the company</title>\n"
            "<!-- Written by scripts/companybook.py from content/company/finance.yaml, company.yaml, people.yaml, fleet.yaml and the snapshot.\n"
            "     Do not edit this page: change the yaml and run .venv/bin/python scripts/build.py -->\n"
            f"<style>{CSS}</style>\n</head>\n<body>\n<div class=\"wrap\">\n" + "\n".join(out)
            + '\n<p class="note" style="margin-top:28px">Written by the build from content/company/finance.yaml. To change a fact, change it there and run the build.</p>\n</div>\n</body>\n</html>\n')
    target = ROOT / "docs" / "company.html"
    target.parent.mkdir(exist_ok=True)
    target.write_text(html, encoding="utf-8")
