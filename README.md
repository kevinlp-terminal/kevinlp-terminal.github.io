# KEVIN LP | UK M&A Terminal

Live site: https://kevinlp-terminal.github.io

UK and US takeover tracker (UK firm offers under the Takeover Code; largest US deals, $1bn+): deal list, premiums, comps, live deal health and price log. Data refreshed each weekday morning from the KEVIN LP workbook.

## How it updates
A GitHub Action (`.github/workflows/refresh.yml`) runs four times every weekday. It fetches the latest prices for every live UK deal (`engine/update_prices.py`) and US deal (`engine/us_update.py`, which also backfills pre-announcement prices), logs them in `engine/UK_MA_Terminal.xlsx`, rebuilds `index.html` and publishes it. No laptop needed. Run it by hand from the Actions tab ("Refresh KEVIN LP" → Run workflow).
