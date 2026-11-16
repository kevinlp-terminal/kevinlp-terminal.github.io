"""Build the KEVIN LP web terminal from the workbook.

Usage:  python build_web.py <UK_MA_Terminal.xlsx> <template.html> <output.html>

Reads only the input cells (deals, prices, deal-health overrides, milestones, watchlist),
so it works whether or not Excel has recalculated the workbook. The web page recomputes
every premium, spread and health status itself.
"""
import json
import re
import sys
from datetime import date, datetime

from openpyxl import load_workbook

DEAL_COLS = {  # Deals tab column -> snapshot key
    1: "id", 2: "target", 3: "bidder", 4: "type", 5: "hq", 6: "sector", 7: "op_start", 8: "date27",
    10: "consideration", 11: "price", 12: "total", 13: "undisturbed", 15: "rep_prem", 18: "rep_vwap3m",
    19: "equity", 20: "ev", 21: "competing", 22: "proposals", 23: "target_fa", 24: "bidder_fa",
    25: "status", 26: "note", 27: "source",
}


def iso(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return v


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def build(wb_path):
    wb = load_workbook(wb_path)  # formulas as text; inputs as values
    ds = wb["Deals"]
    deals, row_id = [], {}
    r = 5
    while ds.cell(r, 2).value not in (None, ""):
        d = {}
        for c, k in DEAL_COLS.items():
            v = ds.cell(r, c).value
            if is_formula(v):
                v = None
            if c == 27 and ds.cell(r, c).hyperlink is not None and ds.cell(r, c).hyperlink.target:
                v = ds.cell(r, c).hyperlink.target
            d[k] = iso(v)
        if d["id"] is None:
            d["id"] = r - 4
        deals.append(d)
        row_id[r] = d["id"]
        r += 1
    known = {d["id"] for d in deals}

    prices = []
    pl = wb["Price Log"]
    for r in range(5, pl.max_row + 1):
        dt, i, px = pl.cell(r, 1).value, pl.cell(r, 2).value, pl.cell(r, 4).value
        if isinstance(dt, (date, datetime)) and num(i) is not None and num(px) is not None and int(i) in known:
            prices.append({"date": iso(dt), "id": int(i), "price": float(px), "note": pl.cell(r, 7).value or ""})

    exp_close, fallback, notes, offer_ov = {}, {}, {}, {}
    dh = wb["Deal Health"]
    for r in range(15, dh.max_row + 1):
        b = dh.cell(r, 2).value
        if not is_formula(b):
            continue
        m = re.search(r"Deals!\$?A\$?(\d+)", b)
        if not m or int(m.group(1)) not in row_id:
            continue
        i = str(row_id[int(m.group(1))])
        e, f, g, s = (dh.cell(r, c).value for c in (5, 6, 7, 19))
        if num(e) is not None:
            offer_ov[i] = e
        if num(f) is not None:
            fallback[i] = f
        if isinstance(g, (date, datetime)):
            exp_close[i] = iso(g)
        if s:
            notes[i] = str(s)

    milestones, watch = {}, []
    lt = wb["Live Tracker"]
    on = False
    for r in range(6, lt.max_row + 1):
        b = lt.cell(r, 2).value
        if isinstance(b, str) and b.startswith("Possible offers watchlist"):
            on = True
            continue
        if not on and is_formula(b):
            m = re.search(r"Deals!\$?B\$?(\d+)", b)
            if m and int(m.group(1)) in row_id:
                tgt = next(d["target"] for d in deals if d["id"] == row_id[int(m.group(1))])
                milestones[tgt] = lt.cell(r, 8).value or ""
        if on and isinstance(b, str) and b and b != "Target":
            watch.append({"target": b, "bidder": lt.cell(r, 3).value or "", "reported": lt.cell(r, 4).value or "",
                          "value": num(lt.cell(r, 5).value), "consideration": lt.cell(r, 6).value or "",
                          "note": lt.cell(r, 8).value or ""})

    sectors = sorted({d["sector"] for d in deals if d["sector"]} | {
        "Business Services", "Consumer", "Energy", "Financials", "Healthcare", "Industrials",
        "Real Estate", "Technology & Telecoms", "Travel & Leisure"})
    for d in deals:
        d["region"], d["ccy"] = "UK", "p"
    add_us(deals, prices, offer_ov)
    latest = max((p["date"] for p in prices), default=None)
    return {"deals": deals, "prices": prices, "exp_close": exp_close, "fallback": fallback,
            "health_notes": notes, "offer_override": offer_ov, "milestones": milestones, "watch": watch,
            "sectors": sectors, "built": datetime.now().strftime("%Y-%m-%d %H:%M"), "latest_price": latest}


def add_us(deals, prices, offer_ov):
    """Merge US deals (us_deals.csv / us_prices.csv next to the workbook's engine folder)."""
    import csv, os
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(base, "us_deals.csv")
    if not os.path.exists(path):
        return
    fl = lambda x: float(x) if x not in (None, "") else None
    for r in csv.DictReader(open(path, newline="", encoding="utf-8")):
        price = fl(r["price"])
        deals.append({"id": int(r["id"]), "target": r["target"], "bidder": r["bidder"], "type": r["type"], "hq": "USA",
                      "sector": r["sector"], "op_start": None, "date27": r["date27"], "consideration": r["consideration"],
                      "price": price, "total": price, "undisturbed": fl(r["undisturbed"]), "rep_prem": fl(r["rep_prem"]),
                      "rep_vwap3m": None, "equity": fl(r["equity"]), "ev": None, "competing": "", "proposals": None,
                      "target_fa": "", "bidder_fa": "", "status": r["status"], "note": r["note"], "source": r["source"],
                      "region": "US", "ccy": "$", "ticker": r["ticker"]})
    pp = os.path.join(base, "us_prices.csv")
    if os.path.exists(pp):
        last = {}
        for p in csv.DictReader(open(pp, newline="", encoding="utf-8")):
            prices.append({"date": p["date"], "id": int(p["id"]), "price": float(p["price"]), "note": p["note"]})
            if p["offer"]:
                last[p["id"]] = (p["date"], float(p["offer"]))
        for i, (_, o) in last.items():
            offer_ov[i] = o


if __name__ == "__main__":
    wb_path, tpl, out = sys.argv[1:4]
    snap = build(wb_path)
    html = open(tpl, encoding="utf-8").read().replace("__SNAPSHOT__", json.dumps(snap).replace("</", "<\\/"))
    open(out, "w", encoding="utf-8").write(html)
    print(f"{len(snap['deals'])} deals, {len(snap['prices'])} prices (latest {snap['latest_price']}), "
          f"{len(snap['watch'])} watchlist -> {out}")
