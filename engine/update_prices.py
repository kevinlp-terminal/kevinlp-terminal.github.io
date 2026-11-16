"""Daily price updater for the UK M&A Terminal.

Fetches the latest closing price for every live (Pending) deal and appends it to the
Price Log tab. Deal Health and the Terminal home screen then update when Excel opens.

Usage (run from this folder):
    python update_prices.py                      # fetch from Yahoo Finance, Stooq as fallback
    python update_prices.py --dry-run            # show what would be written, change nothing
    python update_prices.py --manual 8=425.5 13=702  # type prices in yourself (deal_id=pence)
    python update_prices.py --source stooq       # force one data source

Prices are stored in pence, like the rest of the workbook.
"""
import argparse
import csv
import io
import math
import sys
from datetime import date, datetime
from pathlib import Path

from terminal_tools.common import (CYAN, GREY, LOG_FIRST, backup, black_cell, deals_table,
                                   open_workbook, save_workbook)

HERE = Path(__file__).resolve().parent
DEFAULT_WB = HERE / "UK_MA_Terminal.xlsx"
TICKERS = HERE / "tickers.csv"


# ------------------------------------------------------------------ data sources
def fetch_yahoo(symbol):
    """Return (date, price_in_pence) or None."""
    try:
        import yfinance as yf
    except ImportError:
        return None
    try:
        t = yf.Ticker(symbol)
        hist = t.history(period="7d", auto_adjust=False)
        if hist.empty:
            return None
        last = hist.dropna(subset=["Close"]).iloc[-1]
        px = float(last["Close"])
        cur = None
        try:
            cur = t.fast_info.get("currency")
        except Exception:
            pass
        if cur == "GBP":          # quoted in pounds -> convert to pence
            px *= 100
        return hist.dropna(subset=["Close"]).index[-1].date(), px
    except Exception as e:  # network or symbol problem
        print(f"   yahoo {symbol}: {e}")
        return None


def fetch_stooq(symbol):
    import requests
    url = f"https://stooq.com/q/l/?s={symbol}&f=sd2t2ohlcv&h&e=csv"
    try:
        r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        r.raise_for_status()
        rows = list(csv.DictReader(io.StringIO(r.text)))
        if not rows or rows[0].get("Close") in (None, "", "N/D"):
            return None
        return datetime.strptime(rows[0]["Date"], "%Y-%m-%d").date(), float(rows[0]["Close"])
    except Exception as e:
        print(f"   stooq {symbol}: {e}")
        return None


def sanity(px, offer):
    """Fix pounds/pence mix-ups and reject nonsense: a live target trades near its offer."""
    if not offer or not px:
        return px, "no offer to check against"
    ratio = px / offer
    if 0.2 <= ratio <= 2.0:
        return px, ""
    if 0.2 <= ratio * 100 <= 2.0:
        return px * 100, "converted GBP to pence"
    return None, f"rejected: {px:.2f} is far from offer {offer:.2f}"


# ------------------------------------------------------------------ workbook
def existing_rows(ws):
    seen, r = {}, LOG_FIRST
    first_empty = None
    while r <= ws.max_row:
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a in (None, "") and b in (None, ""):
            if first_empty is None:
                first_empty = r
        else:
            d = a.date() if isinstance(a, datetime) else a
            seen[(d, b)] = r
        r += 1
    return seen, first_empty


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", type=Path, default=DEFAULT_WB)
    ap.add_argument("--source", choices=["auto", "yahoo", "stooq"], default="auto")
    ap.add_argument("--manual", nargs="*", default=[], help="deal_id=price_in_pence pairs")
    ap.add_argument("--date", help="date for manual prices, YYYY-MM-DD (default today)")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--refresh-today", action="store_true",
                    help="overwrite today's automatic price with a fresher one (for hourly runs)")
    args = ap.parse_args()

    wb = open_workbook(args.workbook)
    deals = {d["id"]: d for d in deals_table(wb)}
    live = {i: d for i, d in deals.items() if d["status"] == "Pending"}
    tick = {int(r["deal_id"]): r for r in csv.DictReader(open(TICKERS, newline="", encoding="utf-8"))}

    results = []   # (date, id, price, note)
    if args.manual:
        d0 = datetime.strptime(args.date, "%Y-%m-%d").date() if args.date else date.today()
        for pair in args.manual:
            i, p = pair.split("=")
            results.append((d0, int(i), float(p), "manual"))
    else:
        for i, d in live.items():
            t = tick.get(i)
            if not t:
                print(f"!! {d['target']}: no ticker in tickers.csv - skipped")
                continue
            got, src = None, None
            if args.source in ("auto", "yahoo"):
                got, src = fetch_yahoo(t["yahoo"]), "yahoo"
            if got is None and args.source in ("auto", "stooq"):
                got, src = fetch_stooq(t["stooq"]), "stooq"
            if got is None:
                print(f"!! {d['target']}: no price found ({t['yahoo']} / {t['stooq']})")
                continue
            px, warn = sanity(got[1], d["total"])
            if px is None:
                print(f"!! {d['target']}: {warn}")
                continue
            results.append((got[0], i, round(px, 4), f"auto: {src} close" + (f" ({warn})" if warn else "")))

    ws = wb["Price Log"]
    seen, row = existing_rows(ws)
    written = 0
    print(f"\n{'DATE':<12}{'ID':>4}  {'TARGET':<30}{'PRICE (p)':>11}{'OFFER (p)':>11}{'SPREAD':>9}")
    for d0, i, px, note in results:
        deal = deals.get(i, {})
        offer = deal.get("total")
        spread = f"{offer / px - 1:8.2%}" if offer else "     n/a"
        flag = ""
        if (d0, i) in seen:
            er = seen[(d0, i)]
            old_note = str(ws.cell(er, 7).value or "")
            if args.refresh_today and d0 == date.today() and old_note.startswith("auto") and not args.dry_run:
                if ws.cell(er, 4).value != px:
                    ws.cell(er, 4).value = px
                    ws.cell(er, 7).value = note + f" (refreshed {datetime.now():%H:%M})"
                    written += 1
                    flag = "  (refreshed)"
                else:
                    flag = "  (unchanged)"
            else:
                flag = "  (already logged)"
        print(f"{d0:%d-%b-%y}  {i:>4}  {str(deal.get('target', '?'))[:29]:<30}{px:>11,.2f}{(offer or 0):>11,.2f}{spread}{flag}")
        if flag or args.dry_run:
            continue
        if row is None:
            print("!! Price Log is full - add more rows (copy the last formula row down).")
            break
        for col, val, color, fmt in ((1, d0, CYAN, "dd-mmm-yy"), (2, i, CYAN, None),
                                     (4, px, CYAN, '#,##0.00;(#,##0.00);"-"'), (7, note, GREY, None)):
            black_cell(ws.cell(row, col, val), color, fmt=fmt)
        seen[(d0, i)] = row
        written += 1
        row += 1
        while row <= ws.max_row and ws.cell(row, 1).value not in (None, ""):
            row += 1

    if args.dry_run:
        print("\nDry run - nothing saved.")
        return 0
    if written:
        try:
            b = backup(args.workbook)
            save_workbook(wb, args.workbook)
        except PermissionError:
            print("\n!! The workbook is open in Excel, so nothing was saved. Close it and run again.")
            return 2
        print(f"\nWrote {written} price(s). Backup: {b.name}. Open the workbook in Excel to refresh Deal Health.")
    else:
        print("\nNothing new to write.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
