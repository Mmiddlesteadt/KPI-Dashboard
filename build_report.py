#!/usr/bin/env python3
"""
WTDC monthly customer report builder.

Turns the internal monthly KPI workbook into one self-contained, read-only HTML
report that can be emailed to the customer. The customer only opens the file:
no uploads, no logins, nothing sent anywhere.

Usage
    python build_report.py Campari_October_2026_KPI.xlsx
    python build_report.py Campari_October_2026_KPI.xlsx -o reports/

The script prints a reconciliation check for WTDC staff before you send.
Those checks are for internal use and never appear in the customer's report.
Requires: openpyxl, numpy (pip install openpyxl numpy)
"""
import argparse
import collections
import datetime as dt
import json
import re
import sys
from pathlib import Path

import numpy as np
import openpyxl

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "templates" / "report_template.html"
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]


# ---------------------------------------------------------------- helpers
def clean(v):
    """Strip the non-breaking-space indentation the warehouse export uses."""
    return v.replace("\xa0", " ").strip() if isinstance(v, str) else v


def iso(d):
    if isinstance(d, dt.datetime):
        d = d.date()
    return d.isoformat() if isinstance(d, dt.date) else None


def num(v):
    if v is None or v == "":
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f.is_integer() else f


def networkdays_minus_one(a, b):
    """Same as Excel's =NETWORKDAYS(a,b)-1 (weekdays only, no holiday list)."""
    if not a or not b:
        return None
    a, b = np.datetime64(a.date() if isinstance(a, dt.datetime) else a), \
        np.datetime64(b.date() if isinstance(b, dt.datetime) else b)
    one = np.timedelta64(1, "D")
    n = np.busday_count(a, b + one) if a <= b else -np.busday_count(b, a + one)
    return int(n) - 1


def find_sheet(wb, *names):
    lower = {ws.title.lower().strip(): ws for ws in wb}
    for n in names:
        if n.lower() in lower:
            return lower[n.lower()]
    raise SystemExit(f"Workbook is missing a sheet named one of: {names}")


# ---------------------------------------------------------------- readers
def read_period(ws):
    """'September 1 - 30, 2026' and 'Tuesday, October 13, 2026' from the summary header."""
    period_text, run_text = clean(ws["A3"].value), clean(ws["A4"].value)
    m = re.match(r"([A-Za-z]+)\s+(\d+)\s*-\s*(\d+),\s*(\d{4})", period_text or "")
    if not m:
        raise SystemExit(f"Could not read the reporting period from 'In Out Summary'!A3: {period_text!r}")
    month = MONTHS.index(m.group(1).capitalize()) + 1
    year = int(m.group(4))
    start = dt.date(year, month, int(m.group(2)))
    end = dt.date(year, month, int(m.group(3)))
    run = None
    try:
        run = dt.datetime.strptime(run_text, "%A, %B %d, %Y").date()
    except (TypeError, ValueError):
        pass
    return dict(month=month, year=year, start=start.isoformat(), end=end.isoformat(),
                runDate=run.isoformat() if run else None)


def read_summary(ws):
    header_row = None
    for r in range(1, 15):
        if clean(ws.cell(r, 1).value) == "Part Number":
            header_row = r
            break
    if not header_row:
        raise SystemExit("Could not find the 'Part Number' header on 'In Out Summary'.")
    products, total = [], None
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        part, desc = clean(row[0]), clean(row[1])
        vals = [num(v) for v in row[2:6]]
        if part in (None, "") and all(v is not None for v in vals):
            total = dict(zip(["start", "in", "out", "end"], vals))
            continue
        if part in (None, ""):
            continue
        products.append(dict(part=str(part), desc=desc or "", start=vals[0] or 0,
                             inb=vals[1] or 0, out=vals[2] or 0, end=vals[3] or 0))
    return products, total


def read_detail(ws, descs):
    """Parse the grouped In/Out detail export into receipt lines and shipment lines."""
    ins, outs, customer = [], [], None
    section, part = None, None
    for row in ws.iter_rows(min_row=1, values_only=True):
        a = clean(row[0])
        if not isinstance(a, str) or not a:
            continue
        if re.fullmatch(r"In \(\d+\)", a):
            section = "in"; continue
        if re.fullmatch(r"Out \(\d+\)", a):
            section = "out"; continue
        if a.startswith("Inventory ") or a == "Status":
            continue
        if row[5] is not None and row[7] is not None and section:
            rec = dict(status=a, entryDate=row[1], wr=row[3], po=row[4], part=str(clean(row[5])),
                       pcs=num(row[7]) or 0, release=row[8], consignee=clean(row[9]),
                       releasePo=row[10], outDate=row[11])
            (ins if section == "in" else outs).append(rec)
            continue
        # A product block header: "<part> <description> <customer>"
        first = a.split()[0]
        if first in descs:
            part, section = first, None
            rest = a[len(first):].strip()
            d = descs[first]
            if customer is None and rest.startswith(d):
                customer = rest[len(d):].strip() or None
    return ins, outs, customer


def read_overview(ws, year, month):
    head = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    cols = [(i + 1, h) for i, h in enumerate(head) if isinstance(h, (dt.datetime, dt.date))]
    rows = {}
    for r in range(2, ws.max_row + 1):
        label = clean(ws.cell(r, 1).value)
        if not label:
            continue
        rows[label] = [num(ws.cell(r, c).value) for c, _ in cols]
    months = [f"{h.year}-{h.month:02d}" for _, h in cols]
    # keep months up to and including the report month that have volume data
    target = f"{year}-{month:02d}"
    if target not in months:
        raise SystemExit(f"'Overview' has no column for {target}.")
    keep = [i for i, m in enumerate(months) if m <= target]
    return dict(months=[months[i] for i in keep],
                rows={k: [v[i] for i in keep] for k, v in rows.items()})


def read_comparison(ws):
    out = {}
    for row in ws.iter_rows(min_row=1, values_only=True):
        y = num(row[0])
        if isinstance(y, int) and 1990 < y < 2100:
            out[str(y)] = [num(v) for v in row[1:13]]
    return out


def read_kpi(wb_vals):
    ws = find_sheet(wb_vals, "KPI")
    orders = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] in (None, ""):
            continue
        created, confirmed, picked, delivered = row[4], row[5], row[7], row[9]
        d1 = num(row[6]) if num(row[6]) is not None else networkdays_minus_one(created, confirmed)
        d2 = num(row[8]) if num(row[8]) is not None else networkdays_minus_one(confirmed, picked)
        d3 = num(row[10]) if isinstance(row[10], (int, float)) else networkdays_minus_one(picked, delivered)
        orders.append(dict(release=str(num(row[0])), po=str(num(row[1]) or ""), pro=str(clean(row[2]) or ""),
                           consignee=clean(row[3]), created=iso(created), confirmed=iso(confirmed),
                           picked=iso(picked), delivered=iso(delivered), d1=d1, d2=d2, d3=d3,
                           notes=clean(row[11]) or ""))
    return orders


def read_co2(wb_vals):
    ws = find_sheet(wb_vals, "CO2")
    moves = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] in (None, ""):
            continue
        lbs, miles, factor = num(row[3]), num(row[4]), num(row[5])
        tonnes = num(row[6])
        if tonnes is None and None not in (lbs, miles, factor):
            tonnes = (lbs / 2000) * miles * factor / 1_000_000
        moves.append(dict(release=str(num(row[0])), po=str(num(row[1])), consignee=clean(row[2]),
                          lbs=lbs, miles=miles, factor=factor, tonnes=tonnes))
    return moves


def read_aged(wb_vals, descs):
    ws = find_sheet(wb_vals, "Inventory >1 Year")
    groups = collections.OrderedDict()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if clean(row[0]) in (None, "") or row[4] in (None, ""):
            continue
        part = str(clean(row[4]))
        key = (part, str(num(row[1])), iso(row[3]))
        g = groups.setdefault(key, dict(part=part, desc=descs.get(part, clean(row[5])), wr=key[1],
                                        po=str(clean(row[2])), date=key[2], pcs=0))
        g["pcs"] += num(row[7]) or 0
    return list(groups.values())


def read_accident(wb_vals):
    ws = find_sheet(wb_vals, "Days Accident Free")
    label, days = clean(ws["A1"].value) or "", num(ws["A2"].value)
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", label)
    as_of = dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2))).isoformat() if m else None
    return dict(days=days, asOf=as_of)


# ---------------------------------------------------------------- build
def build(xlsx: Path, out_dir: Path):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    period = read_period(find_sheet(wb, "In Out Summary"))
    products, totals = read_summary(find_sheet(wb, "In Out Summary"))
    descs = {p["part"]: p["desc"] for p in products}
    ins, outs, customer = read_detail(find_sheet(wb, "In Out Detailed"), descs)
    overview = read_overview(find_sheet(wb, "Overview"), period["year"], period["month"])
    inbound_hist = read_comparison(find_sheet(wb, "Inbound Comparison"))
    outbound_hist = read_comparison(find_sheet(wb, "Outbound Comparison"))
    orders = read_kpi(wb)
    co2 = read_co2(wb)
    aged = read_aged(wb, descs)
    accident = read_accident(wb)
    customer = customer or clean(find_sheet(wb, "Overview")["A1"].value)

    # Receipts: one line per product per warehouse receipt (pallet lines summed)
    rec = collections.OrderedDict()
    for r in ins:
        k = (r["part"], str(r["wr"]))
        g = rec.setdefault(k, dict(part=r["part"], wr=str(r["wr"]), po=str(r["po"]), date=iso(r["entryDate"]), pcs=0))
        g["pcs"] += r["pcs"]
    # Shipments: one line per product per shipment (release)
    shp = collections.OrderedDict()
    for r in outs:
        k = (r["part"], str(r["release"]))
        g = shp.setdefault(k, dict(part=r["part"], release=str(r["release"]), consignee=r["consignee"],
                                   po=str(r["releasePo"]), date=iso(r["outDate"]), pcs=0))
        g["pcs"] += r["pcs"]

    data = dict(
        meta=dict(customer=customer, **period, generated=dt.date.today().isoformat()),
        overview=overview,
        products=products,
        totals=totals,
        receipts=list(rec.values()),
        shipments=list(shp.values()),
        history=dict(inbound=inbound_hist, outbound=outbound_hist),
        orders=orders,
        co2=co2,
        aged=aged,
        accident=accident,
    )

    issues = checks(data, ins, outs)

    html = TEMPLATE.read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = html.replace("/*__REPORT_DATA__*/null", payload)
    month_name = MONTHS[period["month"] - 1]
    safe_customer = re.sub(r"[^A-Za-z0-9]+", "-", customer or "Customer").strip("-")
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"WTDC-Monthly-Report-{safe_customer}-{month_name}-{period['year']}.html"
    out.write_text(html, encoding="utf-8")
    return out, issues


def checks(d, ins, outs):
    """Internal reconciliation. Printed for WTDC staff, never shown to the customer."""
    issues, ok = [], []
    P = d["products"]
    s = {k: sum(p[k] for p in P) for k in ("start", "inb", "out", "end")}
    ov = d["overview"]["rows"]
    last = lambda label: (ov.get(label) or [None])[-1]
    m = d["meta"]["month"] - 1
    yr = str(d["meta"]["year"])

    def same(label, a, b):
        (ok if a == b else issues).append(f"{label}: {a:,} vs {b:,}" if isinstance(a, (int, float)) and isinstance(b, (int, float)) else f"{label}: {a} vs {b}")

    bad = [p["part"] for p in P if p["start"] + p["inb"] - p["out"] != p["end"]]
    (issues if bad else ok).append(f"Start + In - Out = End for every product" + (f" (fails: {', '.join(bad)})" if bad else ""))
    if d["totals"]:
        same("Summary total row vs sum of products (end)", d["totals"]["end"], s["end"])
    same("Cases received: summary vs Overview", s["inb"], last("Inbound Cases Volume"))
    same("Cases shipped: summary vs Overview", s["out"], last("Outbound Cases Volume"))
    same("Month-end inventory: summary vs Overview", s["end"], last("Month End Inventory"))
    same("Cases received: detail lines vs summary", sum(r["pcs"] for r in ins), s["inb"])
    same("Cases shipped: detail lines vs summary", sum(r["pcs"] for r in outs), s["out"])
    same("Cases received: Inbound Comparison vs Overview", d["history"]["inbound"].get(yr, [None] * 12)[m], last("Inbound Cases Volume"))
    same("Cases shipped: Outbound Comparison vs Overview", d["history"]["outbound"].get(yr, [None] * 12)[m], last("Outbound Cases Volume"))
    prev_end = (ov.get("Month End Inventory") or [None, None])[-2] if len(ov.get("Month End Inventory", [])) > 1 else None
    if prev_end is not None:
        same("Opening inventory vs last month's closing inventory", s["start"], prev_end)
    releases = {r["release"] for r in d["shipments"]}
    same("Shipments in detail vs \"PO's Shipped\" on Overview", len(releases), last("PO's Shipped"))
    logged = {o["release"] for o in d["orders"]}
    missing = sorted(releases - logged)
    if missing:
        dest = {r["release"]: r["consignee"] for r in d["shipments"]}
        note = ", ".join(f"{x} ({dest[x]})" for x in missing)
        (ok if all(dest[x] == "Destruction" for x in missing) else issues).append(
            f"Shipments not in the KPI delivery log: {note}")
    neg = [o for o in d["orders"] if any((o[k] or 0) < 0 for k in ("d1", "d2", "d3"))]
    for o in neg:
        issues.append(f"KPI order {o['release']}: a date is earlier than the step before it "
                      f"(created {o['created']}, confirmed {o['confirmed']}, picked {o['picked']}, "
                      f"delivered {o['delivered']}). Days by step: {o['d1']}, {o['d2']}, {o['d3']}")
    aged_total = sum(a["pcs"] for a in d["aged"])
    over = [a for a in d["aged"] if a["date"] and a["date"] > (dt.date.fromisoformat(d["meta"]["end"]) - dt.timedelta(days=365)).isoformat()]
    if over:
        issues.append(f"{len(over)} 'Inventory >1 Year' lines were received less than a year before {d['meta']['end']}")
    ok.append(f"Inventory held over one year: {aged_total:,} cases on {len({a['wr'] for a in d['aged']})} receipts")
    return dict(ok=ok, issues=issues)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("-o", "--out", type=Path, default=HERE / "output")
    a = ap.parse_args()
    out, res = build(a.workbook, a.out)
    print(f"\nReport written: {out}\n")
    print("Internal checks (not shown to the customer)")
    for line in res["ok"]:
        print(f"  OK     {line}")
    for line in res["issues"]:
        print(f"  CHECK  {line}")
    if res["issues"]:
        print(f"\n{len(res['issues'])} item(s) to review in the workbook before sending.")
    else:
        print("\nEverything reconciles.")


if __name__ == "__main__":
    main()
