"""US deals: backfill unaffected prices and log daily prices for live deals.

Reads us_deals.csv. For each deal:
  * if 'undisturbed' is blank, fills it with the target's close on the last trading day
    before the announcement date (labelled as such in the page);
  * for share deals with no announcement value, fills 'price' with cash + ratio x the
    acquirer's close on that same day;
  * for Pending deals, logs today's target close and the current implied offer value
    (cash + ratio x acquirer close) in us_prices.csv (one row per deal per day).
Network problems are reported and skipped; nothing is overwritten with blanks.
"""
import csv
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEALS, PRICES = HERE / "us_deals.csv", HERE / "us_prices.csv"


def hist(sym, start, end):
    import yfinance as yf
    try:
        h = yf.Ticker(sym).history(start=start, end=end, auto_adjust=False)
        return h.dropna(subset=["Close"])
    except Exception as e:
        print(f"   {sym}: {e}")
        return None


def close_before(sym, d):
    h = hist(sym, d - timedelta(days=10), d)
    if h is None or h.empty:
        return None
    return round(float(h["Close"].iloc[-1]), 4)


def latest(sym):
    h = hist(sym, date.today() - timedelta(days=10), date.today() + timedelta(days=1))
    if h is None or h.empty:
        return None, None
    return h.index[-1].date(), round(float(h["Close"].iloc[-1]), 4)


def f(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def main():
    rows = list(csv.DictReader(open(DEALS, newline="", encoding="utf-8")))
    fields = list(rows[0].keys())
    changed = False
    for r in rows:
        d27 = datetime.strptime(r["date27"], "%Y-%m-%d").date()
        if not r["undisturbed"] and r["ticker"]:
            u = close_before(r["ticker"], d27)
            if u:
                r["undisturbed"] = str(u)
                r["note"] = (r["note"] + " Undisturbed = close on last trading day before announcement.").strip()
                changed = True
                print(f"{r['target']}: undisturbed {u}")
        if not r["price"] and f(r["ratio"]) and r["acq_ticker"]:
            a = close_before(r["acq_ticker"], d27)
            if a:
                r["price"] = str(round((f(r["cash"]) or 0) + f(r["ratio"]) * a, 4))
                changed = True
                print(f"{r['target']}: offer value at announcement {r['price']}")
    if changed:
        with open(DEALS, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=fields)
            w.writeheader()
            w.writerows(rows)

    log = list(csv.DictReader(open(PRICES, newline="", encoding="utf-8"))) if PRICES.exists() else []
    key = {(p["date"], p["id"]): p for p in log}
    n = 0
    for r in rows:
        if r["status"] != "Pending" or not r["ticker"]:
            continue
        dt, px = latest(r["ticker"])
        if px is None:
            print(f"!! {r['target']}: no price")
            continue
        offer = f(r["price"])
        if f(r["ratio"]) and r["acq_ticker"]:
            _, a = latest(r["acq_ticker"])
            if a:
                offer = round((f(r["cash"]) or 0) + f(r["ratio"]) * a, 4)
        rec = {"date": dt.isoformat(), "id": r["id"], "price": str(px), "offer": "" if offer is None else str(offer),
               "note": "auto: yahoo close"}
        key[(rec["date"], rec["id"])] = rec
        n += 1
    with open(PRICES, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["date", "id", "price", "offer", "note"])
        w.writeheader()
        w.writerows(sorted(key.values(), key=lambda p: (p["date"], int(p["id"]))))
    print(f"US: {n} live prices logged")
    return 0


if __name__ == "__main__":
    sys.exit(main())
