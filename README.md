# KEVIN LP | UK M&A Terminal

Live site: https://kevinlp-terminal.github.io

UK takeover tracker: deal list, premiums, comps, live deal health and price log. Data refreshed each weekday morning from the KEVIN LP workbook.

## How it updates
A GitHub Action (`.github/workflows/refresh.yml`) runs four times every weekday. It fetches the latest prices for every live deal (`engine/update_prices.py`), logs them in `engine/UK_MA_Terminal.xlsx`, rebuilds `index.html` and publishes it. No laptop needed. Run it by hand from the Actions tab ("Refresh KEVIN LP" → Run workflow).
