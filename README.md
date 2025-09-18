
# Mongolian News Ad Scraper (ikon.mn, gogo.mn, news.mn)

This project scrapes homepage/banner creatives from three Mongolian news sites and attributes them to advertisers:
- **gogo.mn** (Boost carousel)
- **ikon.mn** (`/ad/` pages)
- **news.mn** (strict banner container)

It saves each creative to disk, appends rows to a CSV, and maintains a persistent **banner ledger** that de-duplicates by MD5 and pHash and aggregates advertiser attribution across days.

## Quick start

```bash
python -m venv .venv
# Windows:
. .venv/Scripts/activate
# macOS/Linux:
# source .venv/bin/activate

pip install -r requirements.txt
python - <<'PY'
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    pass
PY
playwright install

# Run all scrapers
python run.py

# Or just one site:
python run.py --gogo
python run.py --ikon
python run.py --news
```

### CLI flags

```
python run.py --help

--output <dir>           Root folder for saved banners (default: C:/Data/work/webbot/banner_screenshots)
--csv <file>             CSV file to append combined results (default: C:/Data/work/webbot/banner_tracking_combined.csv)
--ledger <file>          Path to banner ledger CSV (default: ./banner_master.csv)
--no-skip-gifs           Save GIFs too (default: GIFs skipped)
--ikon/--gogo/--news     Selectively run scrapers (default: run all)
--max-mins N             Hard cap minutes per site (default: 6)
--req-timeout-ms N       Per-image download timeout in ms (default: 10000)
```

### Notes

- `banner_ledger.py` keeps a persistent CSV of unique creatives with de-duplication and attribution.
- The **gogo** scraper has a robust resolver chain and optional redirect following to resolve final advertiser hosts.
- If you want verbose debugging, open the file (e.g., `gogo_mn.py`) and set `DEBUG_DETECT = True`.
