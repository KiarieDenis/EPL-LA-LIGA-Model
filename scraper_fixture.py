from seleniumbase import SB
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO
import sys
import os
import shutil

# Configure terminal output for Unicode support on Windows
sys.stdout.reconfigure(encoding='utf-8')

# Output folder for scraped fixture CSVs
OUTPUT_DIR = "FIXTURES COMBINED DATA"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Folder for raw data cleanup – removed old unused directories for a fresh start
for old_dir in ["combined_data", "cleaned_data"]:
    if os.path.isdir(old_dir):
        try:
            shutil.rmtree(old_dir)
            print(f"Removed old directory: {old_dir}")
        except Exception as e:
            print(f"Failed to remove {old_dir}: {e}")

# League configurations for fixtures
LEAGUES = [
    {
        "name": "Premier League (EPL)",
        "url": "https://fbref.com/en/comps/9/schedule/Premier-League-Scores-and-Fixtures",
        "prefix": "premier_league",
    },
    {
        "name": "La Liga",
        "url": "https://fbref.com/en/comps/12/schedule/La-Liga-Scores-and-Fixtures",
        "prefix": "laliga",
    },
]

def scrape_fixtures(league: dict) -> None:
    """Scrape the fixtures & scores table for a given league and write to CSV."""
    url = league["url"]
    prefix = league["prefix"]
    name = league["name"]
    print(f"\n{'='*60}")
    print(f"  Scraping fixtures: {name}")
    print(f"  URL: {url}")
    print(f"{'='*60}")

    with SB(uc=True, test=True) as sb:
        sb.uc_open_with_reconnect(url, 6)
        # Attempt to clear any Cloudflare challenge
        try:
            sb.uc_gui_click_captcha()
        except Exception:
            pass
        # Extract page source and unwrap hidden tables
        html = sb.get_page_source()
        # FBref hides some tables inside HTML comments – remove comment markers
        html = html.replace('<!--', '').replace('-->', '')
        soup = BeautifulSoup(html, "html.parser")

    # Parse all tables and pick the one that looks like a fixture table
    tables = pd.read_html(StringIO(str(soup)))
    df = None
    for tbl in tables:
        cols = [str(c).strip() for c in tbl.columns]
        # Heuristic: must contain a Date column and at least one column with Home/Away/Team info
        if any('date' in c.lower() for c in cols) and (
            any('home' in c.lower() for c in cols) or any('away' in c.lower() for c in cols) or any('team' in c.lower() for c in cols)
        ):
            df = tbl
            # Clean the extracted table: drop completely empty rows and any duplicated header rows that may appear within the data
            df = df.dropna(how='all')
            # Remove rows that are identical to the header (duplicate header rows)
            header = list(df.columns)
            df = df[~df.apply(lambda r: list(r) == header, axis=1)]
            break
    if df is None:
        print(f"  ⚠️  Could not locate fixtures table for {name}")
        print("    Available table columns were:")
        for i, tbl in enumerate(tables):
            print(f"      Table {i}: {list(tbl.columns)}")
        return
    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    csv_path = os.path.join(OUTPUT_DIR, f"{prefix}_2026-2027_fixtures.csv")
    xlsx_path = os.path.join(OUTPUT_DIR, f"{prefix}_2026-2027_fixtures.xlsx")
    df.to_csv(csv_path, index=False)
    # Also write a well‑formatted Excel file
    df.to_excel(xlsx_path, index=False, sheet_name="Fixtures")
    print(f"  ✅ Saved fixtures CSV → {csv_path}")
    print(f"  ✅ Saved fixtures Excel → {xlsx_path}")

# Main execution loop
for league in LEAGUES:
    scrape_fixtures(league)

print("\n✅ All fixture tables scraped successfully!")
