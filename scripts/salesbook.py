"""What Sales keeps about its customers: the contract register, the charter orders, the
customer stories, the account plans, the tracker of each quarter and the charter revenue
by quarter.

Written from `content/company/crm/`: `deals.yaml` and `accounts.yaml` (the CRM),
`contracts.yaml` (terms and charter orders), `account-plans.yaml` and
`case-studies.yaml`. `documents.build()` calls `build()`. The rules of `documents.py`
apply, and three more:

- The trackers are written for every quarter except the third of 2099. That one is the
  win-loss pack's, with its planted error, and nothing here counts its wins and losses.
- A won deal of that quarter is a signed contract like any other, so it is in the
  register and in the revenue by quarter. Its loss reasons and notes are not used.
- Every charter order adds up to the amount of its deal, and every quarter's contracts
  stay below the charter revenue of `finance.yaml`. The build fails otherwise.
"""

from __future__ import annotations

import re
from datetime import date, timedelta
from pathlib import Path

import yaml

import office
from office import AMBER, OK, Book, Doc, F

ROOT = Path(__file__).resolve().parent.parent
DRIVE = ROOT / "content" / "generated" / "drive"
COMPANY = ROOT / "content" / "company"
PACK_DEALS = ROOT / "content" / "packs" / "win-loss" / "crm" / "deals.yaml"
HAND_PLANS = COMPANY / "drive" / "sales" / "accounts"
MONTH = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
ABOUT = "Six ships, ten scheduled routes from Earth to Pluto, and charter rotations built around the customer's calendar. sales@galaxium.example"


def load(name: str):
    return yaml.safe_load((COMPANY / "crm" / name).read_text(encoding="utf-8"))


def long(d) -> str:
    d = date.fromisoformat(str(d)[:10])
    return f"{d.day} {MONTH[d.month]} {d.year}"


def month(ym: str) -> str:
    return f"{MONTH[int(ym[5:7])]} {ym[:4]}"


def months(first: str, last: str) -> list[str]:
    y, m, out = int(first[:4]), int(first[5:7]), []
    while f"{y}-{m:02d}" <= last:
        out.append(f"{y}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def last_day(ym: str) -> date:
    y, m = int(ym[:4]), int(ym[5:7])
    return (date(y + 1, 1, 1) if m == 12 else date(y, m + 1, 1)) - timedelta(days=1)


def quarter(ym: str) -> str:
    return f"{ym[:4]} Q{(int(ym[5:7]) - 1) // 3 + 1}"


def slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


class Sales:
    def __init__(self, today: str) -> None:
        self.today = today
        self.accounts = {a["id"]: a for a in load("accounts.yaml")}
        self.deals = load("deals.yaml")
        self.pack = yaml.safe_load(PACK_DEALS.read_text(encoding="utf-8"))
        book = load("contracts.yaml")
        self.terms, self.orders = book["terms"], {o["deal"]: o for o in book["documents"]}
        won = {x["id"]: x for x in self.deals + self.pack if x["stage"] == "won"}
        assert set(won) == set(self.terms), f"contracts.yaml: terms and won deals differ: {sorted(set(won) ^ set(self.terms))}"
        self.contracts = []
        for did, deal in sorted(won.items(), key=lambda kv: (str(kv[1]["close_date"]), kv[0])):
            first, last = self.terms[did]
            assert str(deal["close_date"])[:7] <= last, f"{did}: signed after its term"
            account = self.accounts[deal["account_id"]]
            file = "vesta-shipwrights-charter-order-2100" if did == "D-2019" else self.orders.get(did, {}).get("file", "")
            self.contracts.append({"id": "GC-" + did[2:], "deal": deal, "account": account, "first": first, "last": last, "months": months(first, last),
                                   "what": deal["name"].split(", ", 1)[-1], "file": f"sales/contracts/{file}.docx" if file else ""})

    def status(self, c: dict, as_of: str) -> str:
        return "Ended" if c["last"] < as_of[:7] else "Signed, not started" if c["first"] > as_of[:7] else "Running"

    def by_quarter(self, c: dict) -> dict[str, int]:
        """The contract's value by quarter of travel: evenly over the months of its term, in whole dollars."""
        amount, n, out = c["deal"]["amount"], len(c["months"]), {}
        for i, ym in enumerate(c["months"]):
            out[quarter(ym)] = out.get(quarter(ym), 0) + amount // n + (amount % n if i == n - 1 else 0)
        return out


# ------------------------------------------------------------------ the register

def register(s: Sales) -> None:
    rows = [{"contract": c["id"], "deal": c["deal"]["id"], "account": c["account"]["name"], "segment": c["account"]["segment"], "what": c["what"], "signed_by": c["deal"]["owner"],
             "signed": str(c["deal"]["close_date"]), "first_month": c["first"], "last_month": c["last"], "months": len(c["months"]), "value": c["deal"]["amount"],
             "value_per_month": F("=K{r}/J{r}", round(c["deal"]["amount"] / len(c["months"]))), "status": s.status(c, s.today), "document": c["file"]} for c in s.contracts]
    book = Book(DRIVE / "sales/contracts/contract-register.xlsx", owner="Sinead Byrne", modified="2099-11-02")
    book.sheet("Contracts", rows, title="Customer contracts signed since January 2098",
               subtitle="One row per won deal in the CRM. Kept by Sales operations, read on Monday 2 November 2099. Status at 4 November 2099.", total={"value": "sum"},
               widths={"account": 32, "what": 44, "document": 62, "segment": 24}, status={"status": {"Running": OK, "Signed, not started": AMBER}})
    running = [c for c in s.contracts if s.status(c, s.today) != "Ended"]
    ends = sorted({c["last"] for c in running})
    book.sheet("Ends by month", [{"last_month": m, "contracts": sum(c["last"] == m for c in running), "value": sum(c["deal"]["amount"] for c in running if c["last"] == m),
                                  "accounts": ", ".join(sorted({c["account"]["name"] for c in running if c["last"] == m}))} for m in ends],
               title="When the running and the signed contracts end", subtitle="By the last month of travel. A contract that ends is a renewal to prepare three months before.",
               total={"contracts": "sum", "value": "sum"}, widths={"accounts": 90})
    book.notes("Notes", ["A contract has the number of its deal: deal D-1744 is contract GC-1744. The CRM is the record for the amount and the signing date.",
                         "The term is the first and the last month of travel. The value is the amount of the deal: base fares at the rate card less the discount, plus what the order lists besides.",
                         "A document path is given where the signed charter order is in sales/contracts. Orders signed before 2098 and the master agreements are kept by Legal.",
                         "Contract flying is never in the booking system. Whether a seat under a contract was used is in the passage lists of operations control."])
    book.close()


# ------------------------------------------------------------------ charter orders and the master terms

def orders(s: Sales) -> None:
    for c in s.contracts:
        o = s.orders.get(c["deal"]["id"])
        if not o:
            continue
        deal, account, signed = c["deal"], c["account"], str(c["deal"]["close_date"])
        term = f"1 {month(c['first'])} to {last_day(c['last']).day} {month(c['last'])}"
        kind = o.get("kind", "Charter order")
        doc = Doc(DRIVE / c["file"], kind, subtitle=f"{account['name']} · {c['what'][0].upper() + c['what'][1:]}", label="Customer contract · Confidential",
                  meta=[("Customer", f"{account['name']}, {account['home']}"), ("Customer contact", ", ".join(o["signer"])), ("Account executive", deal["owner"]),
                        ("Contract", f"{c['id']} (deal {deal['id']})"), ("Signed", long(signed)), ("Term", term)], owner=deal["owner"], modified=signed)
        doc.h1("1. What we fly")
        doc.p(o["flying"])
        doc.h1("2. Price and payment")
        table, fares = [], 0
        for seats, text, rate in o.get("lines", []):
            table.append([f"{text}, at {rate:,}" if seats == 1 and text.startswith("Whole ship") else f"{seats} seats, {text}, at {rate:,}", seats * rate])
            fares += seats * rate
        discount = round(fares * o.get("discount", 0) / 100)
        if discount:
            table.append([f"Discount of {o['discount']}% on the base fares", -discount])
        extras = o.get("extras", [])
        table += [[text, amount] for text, amount in extras]
        total = fares - discount + sum(amount for _, amount in extras)
        assert total == deal["amount"], f"contracts.yaml: the order for {deal['id']} adds up to {total:,}, the deal is {deal['amount']:,}"
        doc.table(["Item", "Dollars"], table + [["Contract value", total]], widths=[4, 1.5], total=True)
        doc.p(f"Payment: {o['payment']}")
        if discount:
            who = "the Head of Sales" if o["discount"] > 5 else "the account executive"
            doc.p(f"The price is the rate card price less {o['discount']}%. {o['discount_why']} The discount was approved by {who} (policies/discount-policy.md) and applies to the base fare only.")
        elif fares:
            doc.p("The price is the rate card price. No discount applies.")
        doc.h1("3. Names and changes")
        doc.bullets(["Names are supplied at booking and must match the travel document.",
                     "A change of traveller is a cancellation and a new booking. Inside the contract it carries no fee up to 72 hours before departure.",
                     "Cancelled seats are returned as travel credit valid for twelve months, not as cash."])
        doc.h1("4. Particular terms")
        doc.bullets(o["special"])
        doc.h1("5. General terms")
        doc.p("The Galaxium master terms for corporate customers apply (legal/templates/master-terms-corporate-customers.docx). Seats under this order fly on contract departures, "
              "on the Luna shuttle or on legs Galaxium flies anyway. They are never taken out of the public seats of a scheduled departure.")
        doc.signatures([("For Galaxium Travels", "Amara Okafor", "Head of Sales", f"Signed {long(signed)}"), (f"For {account['name']}", o["signer"][0], o["signer"][1], f"Signed {long(signed)}")])
        doc.save()

    doc = Doc(DRIVE / "legal/templates/master-terms-corporate-customers.docx", "Master terms for corporate customers", subtitle="The terms behind every charter order", label="Legal · Version 4 · In force from 1 January 2098",
              meta=[("Owner", "Legal"), ("Applies to", "Every charter order and framework order"), ("Replaces", "Version 3 of 2096"), ("Changes need", "The General Counsel")], owner="Helena Vogt", modified="2097-12-06")
    doc.clause("1", "What these terms are", ["These terms apply to every order a corporate customer signs with Galaxium Travels Ltd. An order names what is flown, the price and the term. Where an order and these terms differ, the order applies.",
                                           "A customer with its own master agreement is bound by that agreement in place of these terms."])
    doc.clause("2", "What an order buys", ["A seat or a whole ship on a contract departure, a seat on the Luna shuttle, or a reserved seat on a leg Galaxium flies anyway.",
                                         "An order never buys a public seat of a scheduled departure. The customer's travellers may book those on the public site like anyone else, at the public price.",
                                         "A departure is confirmed only once Galaxium's Operations has confirmed it in writing."])
    doc.clause("3", "Price", ["Prices are those of the rate card in force on the day the order is signed, less the discount the order names.",
                              "A discount applies to the base fare. It does not apply to fuel surcharges, to port fees or to cargo.",
                              "Port fees at a customer's own dock are charged at cost."])
    doc.clause("4", "Names, changes and cancellations", ["Names are supplied at booking and must match the travel document.",
                                                        "A change of traveller is a cancellation and a new booking. It carries no fee up to 72 hours before departure.",
                                                        "A cancelled seat is returned as travel credit valid for twelve months. It is not returned as cash."])
    doc.clause("5", "The way home", ["Where the rate card says that the way home is included, the traveller rides back on a leg Galaxium flies anyway. The seat is asked for through Support at least 48 hours ahead and confirmed when it is free.",
                                    "A customer whose travellers did not fly out with Galaxium may buy the way home on its own, as a reserved seat at half the outbound rate."])
    doc.clause("6", "Changes to the schedule", ["Galaxium gives 14 days' written notice of a change to a departure named in an order.",
                                               "If Galaxium cancels a departure, the customer may take the next departure or have the fee for the affected seats returned."])
    doc.clause("7", "Payment", ["Invoices are payable 30 days net unless the order says otherwise.", "A discount above 5% is given against payment by quarter in advance."])
    doc.clause("8", "Liability", ["Galaxium's liability under an order is limited to the fees for the affected bookings."])
    doc.clause("9", "Term and end", ["An order ends on the last day of its term. It does not renew by itself.", "Either party may end an order with 60 days' written notice if the other is in breach and does not put it right within 30 days."])
    doc.save()


# ------------------------------------------------------------------ customer stories

def stories(s: Sales) -> None:
    for c in load("case-studies.yaml"):
        account = s.accounts[c["account"]]
        doc = Doc(DRIVE / f"sales/case-studies/{c['file']}.docx", c["title"], subtitle=f"{account['name']} and Galaxium Travels", label="Customer story · Approved for external use",
                  meta=[("Customer", f"{account['name']}, {account['home']}"), ("Segment", account["segment"]), ("Customer since", str(account["since"])), ("What we fly", c["flies"])],
                  owner=c["owner"], modified=str(c["modified"]))
        doc.h1("The need")
        doc.p(c["need"])
        doc.h1("What we built")
        doc.bullets(c["built"])
        doc.callout(f"\"{c['quote']}\"  {c['quote_by']}", "In the customer's words")
        doc.h1(c["next_title"])
        doc.p(c["next"])
        doc.h1("About Galaxium Travels")
        doc.p(ABOUT)
        doc.save()


# ------------------------------------------------------------------ account plans

def plans(s: Sales) -> None:
    plan = load("account-plans.yaml")
    for aid, account in s.accounts.items():
        name = f"account-plan-{slug(account['name'])}.md"
        if (HAND_PLANS / name).exists():
            assert aid not in plan, f"{aid} has a hand-written plan and one in account-plans.yaml"
            continue
        p, as_of = plan[aid], str(plan[aid]["modified"])
        since = f"Customer since {account['since']}" if account.get("since") else "Not yet a customer"
        out = ["---", f"owner: {p['by']}", f"modified: {as_of}", "---", f"# Account plan: {account['name']} ({aid})", "",
               f"{account['segment']} · {account['home']} · {account['tier']} · {since} · Account owner: {account['owner']} · Main contact: {account['contact']}", "",
               f"**Who they are.** {p['about']}", "", f"**How they travel.** {p['travel']}", "", f"**Health: {p['health'][0].lower()}.** {p['health'][1]}", "",
               f"**What the customer says.** {p['says']}", "", "## Contracts", ""]
        signed = [c for c in s.contracts if c["account"]["id"] == aid and str(c["deal"]["close_date"]) <= as_of]
        if signed:
            out += ["| Contract | What | Signed | Travel | Value | Status |", "| --- | --- | --- | --- | ---: | --- |"]
            out += [f"| {c['id']} | {c['what'][0].upper() + c['what'][1:]} | {long(c['deal']['close_date'])} | {month(c['first'])} to {month(c['last'])} | {c['deal']['amount']:,} | {s.status(c, as_of)} |" for c in signed]
            out += ["", f"From sales/contracts/contract-register.xlsx. Status at {long(as_of)}. Value signed since January 2098: ${sum(c['deal']['amount'] for c in signed):,}."]
        else:
            out += ["No contract signed since January 2098."]
        open_deals = [x for x in s.deals if x["account_id"] == aid and x["stage"] not in ("won", "lost")]
        if open_deals:
            out += ["", "## Open deals", "", "| Deal | Name | Stage | Close date | Value |", "| --- | --- | --- | --- | ---: |"]
            out += [f"| {x['id']} | {x['name'].split(', ', 1)[-1][0].upper() + x['name'].split(', ', 1)[-1][1:]} | {x['stage'].capitalize()} | {long(x['close_date'])} | {x['amount']:,} |" for x in open_deals]
            out += ["", "The CRM is the record for these."]
        out += ["", "## Plan", ""] + [f"{i}. {step}" for i, step in enumerate(p["plan"], 1)]
        target = DRIVE / "sales" / "accounts" / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("\n".join(out) + "\n", encoding="utf-8")
    assert set(plan) <= set(s.accounts), "account-plans.yaml names an account that is not in accounts.yaml"


# ------------------------------------------------------------------ the tracker of each quarter

def trackers(s: Sales) -> None:
    """Dev Malhotra's tracker, one workbook a quarter, in the columns of the pack's third-quarter tracker."""
    closed = [x for x in s.deals if x["stage"] in ("won", "lost")]
    assert not any("2099-07-01" <= str(x["close_date"]) <= "2099-09-30" for x in closed), "a deal of the third quarter of 2099 belongs in the win-loss pack"
    for q in sorted({quarter(str(x["close_date"])) for x in closed}):
        year, n = q.split(" Q")
        rows = [{"Deal": x["id"], "Name": x["name"], "Account": s.accounts[x["account_id"]]["name"], "Segment": s.accounts[x["account_id"]]["segment"], "Owner": x["owner"],
                 "Amount": x["amount"], "Closed": x["close_date"], "Outcome": x["stage"].capitalize(), "Loss reason": x.get("loss_reason", ""), "Competitor": x.get("competitor", "")}
                for x in sorted(closed, key=lambda x: (str(x["close_date"]), x["id"])) if quarter(str(x["close_date"])) == q]
        end = last_day(f"{year}-{int(n) * 3:02d}")
        friday = end + timedelta(days=(4 - end.weekday()) % 7)
        office.table_book(DRIVE / f"sales/pipeline-tracker-q{n}-{year}.xlsx", {f"Closed deals Q{n} {year}": rows}, plain=True, owner="Dev Malhotra",
                          modified=min(str(friday), "2099-10-30"))


# ------------------------------------------------------------------ charter revenue by quarter

def revenue(s: Sales) -> None:
    finance = yaml.safe_load((COMPANY / "finance.yaml").read_text(encoding="utf-8"))
    year_total = {y["year"]: y["revenue"][1] for y in finance["years"]}
    quarters, total, forecast = [], {}, set()
    for y in finance["charter_by_quarter"]:
        assert sum(y["quarters"]) == year_total[y["year"]], f"finance.yaml: the quarters of {y['year']} do not add up to the year's charter revenue"
        for i, millions in enumerate(y["quarters"], 1):
            quarters.append(f"{y['year']} Q{i}")
            total[quarters[-1]] = millions * 1_000_000
            if i >= y.get("forecast_from", 5):
                forecast.add(quarters[-1])
    split = {c["id"]: s.by_quarter(c) for c in s.contracts}
    from_contracts = {q: sum(v.get(q, 0) for v in split.values()) for q in quarters}
    for q in quarters:
        assert from_contracts[q] <= total[q], f"{q}: the contracts of the register ({from_contracts[q]:,}) are more than the charter revenue of finance.yaml ({total[q]:,})"
    other = "Contracts signed before 2098 and single charters"
    col = lambda q: q.replace(" ", "_")  # noqa: E731
    book = Book(DRIVE / "finance/analysis/charter-revenue-by-quarter.xlsx", owner="Tomás Rivera", modified="2099-11-02")
    book.sheet("By quarter", [{"quarter": q, "basis": "Forecast at the August close" if q in forecast else "Actual", "charter_revenue": total[q], "from_the_contract_register": from_contracts[q],
                               "before_2098_and_single_charters": F("=C{r}-D{r}", total[q] - from_contracts[q])} for q in quarters],
               title="Charter revenue by quarter, 2098 and 2099", subtitle="Dollars. Charter revenue as in finance/annual/financial-summary-2094-2099.xlsx; the split by contract from sales/contracts/contract-register.xlsx.",
               total={"charter_revenue": "sum", "from_the_contract_register": "sum", "before_2098_and_single_charters": "sum"}, bars=["charter_revenue"],
               widths={"basis": 28, "charter_revenue": 18, "from_the_contract_register": 26, "before_2098_and_single_charters": 30})
    names = sorted({c["account"]["name"] for c in s.contracts if any(q in split[c["id"]] for q in quarters)})
    rows = []
    for name in names:
        mine = [c for c in s.contracts if c["account"]["name"] == name]
        rows.append({"account": name, "segment": mine[0]["account"]["segment"], **{col(q): sum(split[c["id"]].get(q, 0) for c in mine) for q in quarters}})
    rows.append({"account": other, "segment": "", **{col(q): total[q] - from_contracts[q] for q in quarters}})
    for i, r in enumerate(rows):
        r["total"] = F(f"=SUM(C{{r}}:{chr(ord('C') + len(quarters) - 1)}{{r}})", sum(r[col(q)] for q in quarters))
    book.sheet("By account", rows, title="Charter revenue by account and quarter", subtitle="Dollars. The column totals are the charter revenue of each quarter.",
               total={**{col(q): "sum" for q in quarters}, "total": "sum"}, widths={"account": 46, "segment": 24})
    book.sheet("By contract", [{"contract": c["id"], "account": c["account"]["name"], "what": c["what"], **{col(q): split[c["id"]].get(q, 0) for q in quarters}}
                               for c in s.contracts if any(q in split[c["id"]] for q in quarters)],
               title="Contract value by quarter of travel", subtitle="Dollars. Contracts of the register with travel in 2098 or 2099.", total={col(q): "sum" for q in quarters}, widths={"account": 32, "what": 44})
    book.notes("Method", ["Charter revenue is recognised in the quarter the travel is flown. The year totals are those of the financial summary.",
                          "A contract of the register is spread evenly over the months of its term. That is how Finance accrues a contract until its trips are flown, and it is close enough for a quarter.",
                          f"\"{other}\" is the rest: contracts signed before the register begins in January 2098, and whole ships chartered once, outside a contract.",
                          "The fourth quarter of 2099 is the forecast made at the August close. The contracts in it are those signed up to 30 October 2099.",
                          "Contracts signed in 2099 with travel in 2100 are in the register and not in this workbook."])
    book.close()


def build(today: str) -> None:
    s = Sales(today)
    register(s)
    orders(s)
    stories(s)
    plans(s)
    trackers(s)
    revenue(s)
