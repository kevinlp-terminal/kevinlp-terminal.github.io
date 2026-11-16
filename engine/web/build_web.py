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
    latest = max((p["date"] for p in prices), default=None)
    return {"deals": deals, "prices": prices, "exp_close": exp_close, "fallback": fallback,
            "health_notes": notes, "offer_override": offer_ov, "milestones": milestones, "watch": watch,
            "sectors": sectors, "built": datetime.now().strftime("%Y-%m-%d %H:%M"), "latest_price": latest}


if __name__ == "__main__":
    wb_path, tpl, out = sys.argv[1:4]
    snap = build(wb_path)
    html = open(tpl, encoding="utf-8").read().replace("__SNAPSHOT__", json.dumps(snap).replace("</", "<\\/"))
    open(out, "w", encoding="utf-8").write(html)
    print(f"{len(snap['deals'])} deals, {len(snap['prices'])} prices (latest {snap['latest_price']}), "
          f"{len(snap['watch'])} watchlist -> {out}")
