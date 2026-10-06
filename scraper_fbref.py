from seleniumbase import SB
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO
import sys
import os
import shutil

# Configure terminal output to support unicode characters on Windows
sys.stdout.reconfigure(encoding='utf-8')

# Folder for final combined Excel files
LEAGUES_DATA_DIR = "LEAGUES COMBINED DATA"
os.makedirs(LEAGUES_DATA_DIR, exist_ok=True)

# Clean up old unused directories for a fresh start
for old_dir in ["combined_data", "cleaned_data", "scraped_data"]:
    if os.path.isdir(old_dir):
        try:
            shutil.rmtree(old_dir)
            print(f"Removed old directory: {old_dir}")
        except Exception as e:
            print(f"Failed to remove {old_dir}: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# League configurations
# Each entry defines:
#   url    – FBref season stats page for the league
#   prefix – file-name prefix used for all output CSVs
#   name   – human-readable name printed in logs
# ─────────────────────────────────────────────────────────────────────────────
LEAGUES = [
    {
        "name":   "Premier League (EPL)",
        "url":    "https://fbref.com/en/comps/9/Premier-League-Stats",
        "prefix": "premier_league",
    },
    {
        "name":   "La Liga",
        "url":    "https://fbref.com/en/comps/12/La-Liga-Stats",
        "prefix": "laliga",
    },
]

# ─────────────────────────────────────────────────────────────────────────────
# Tables to extract from each league page
# suffix – unique suffix of the HTML <table id="…"> attribute on FBref
# name   – human-readable label used in log output
# file   – output filename TEMPLATE  (the league prefix is prepended at runtime)
# ─────────────────────────────────────────────────────────────────────────────
TABLE_TARGETS = [
    {"suffix": "_overall",         "name": "Overall Standings",      "file": "standings.csv"},
    {"suffix": "_home_away",       "name": "Home/Away Standings",    "file": "home_away.csv"},
    {"suffix": "standard_for",     "name": "Squad Standard Stats",   "file": "standard_stats.csv"},
    {"suffix": "keeper_for",       "name": "Squad Goalkeeping",      "file": "goalkeeping.csv"},
    {"suffix": "shooting_for",     "name": "Squad Shooting",         "file": "shooting.csv"},
    {"suffix": "playing_time_for", "name": "Squad Playing Time",     "file": "playing_time.csv"},
    {"suffix": "misc_for",         "name": "Squad Miscellaneous Stats", "file": "misc_stats.csv"},
]


def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Collapse a MultiIndex column header into a single flat string per column."""
    if isinstance(df.columns, pd.MultiIndex):
        new_cols = []
        for col in df.columns:
            parts = [p for p in col if not p.startswith("Unnamed:")]
            new_cols.append("_".join(parts) if parts else "")
        df.columns = new_cols
    return df


def scrape_league(league: dict) -> None:
    """Open a league page with SeleniumBase, parse all target tables, and save in-memory."""
    url    = league["url"]
    prefix = league["prefix"]
    name   = league["name"]

    print(f"\n{'='*60}")
    print(f"  Scraping: {name}")
    print(f"  URL     : {url}")
    print(f"{'='*60}")

    print("Launching Undetected Chromedriver...")
    with SB(uc=True, test=True) as sb:
        print(f"Opening {url}...")
        sb.uc_open_with_reconnect(url, 6)

        print("Checking for Cloudflare Turnstile/CAPTCHA...")
        try:
            sb.uc_gui_click_captcha()
        except Exception:
            pass

        print("Extracting page source...")
        html = sb.get_page_source()

    print("Parsing page source...")
    # FBref hides some tables inside HTML comments – unwrap them before parsing
    html = html.replace("<!--", "").replace("-->", "")
    soup = BeautifulSoup(html, "html.parser")

    all_tables = soup.find_all("table")

    scraped_dfs = {}

    for target in TABLE_TARGETS:
        suffix   = target["suffix"]
        tbl_name = target["name"]

        # Locate the table whose id ends with the expected suffix
        table = next(
            (t for t in all_tables if t.get("id", "").endswith(suffix)),
            None,
        )

        if table is not None:
            df = pd.read_html(StringIO(str(table)))[0]
            df = flatten_columns(df)

            print(f"\n  [{name}] {tbl_name} → {df.shape[0]} rows × {df.shape[1]} cols")
            print(df.head(3).to_string(index=False))

            # Store in-memory
            key = target["file"].replace(".csv", "")
            scraped_dfs[key] = df
        else:
            print(f"  ⚠️  Could not find table '{tbl_name}' (suffix: '{suffix}')")

    # After scraping all tables for this league, combine them into a single Excel file
    combine_league(league, scraped_dfs)


# Table suffixes for combining (must match the CSV filenames without extension)
TABLE_SUFFIXES = [
    "standings",
    "home_away",
    "standard_stats",
    "goalkeeping",
    "shooting",
    "playing_time",
    "misc_stats",
]


def combine_league(league: dict, scraped_dfs: dict) -> None:
    """Process the league‑specific DataFrames from memory, merge them, and write a combined Excel file."""
    prefix = league["prefix"]
    name   = league["name"]
    output_path = os.path.join(LEAGUES_DATA_DIR, f"{prefix}_2026-2027_combined.xlsx")

    print(f"\n{'='*60}")
    print(f"  Combining: {name}")
    print(f"{'='*60}")

    # Load each DataFrame from our in-memory dict
    dfs = {}
    for suffix in TABLE_SUFFIXES:
        if suffix not in scraped_dfs:
            print(f"  ⚠️  Missing data for {suffix} – skipping.")
            continue
        df = scraped_dfs[suffix].copy()
        df.columns = [c.strip() for c in df.columns]
        # Standardise team column name
        if "Squad" in df.columns:
            df = df.rename(columns={"Squad": "Team"})
        else:
            first_col = df.columns[0]
            if df[first_col].dtype == object:
                df = df.rename(columns={first_col: "Team"})
        # Append suffix to columns (except Team) to avoid collisions
        col_suffix = f"_{suffix}" if suffix != "standings" else ""
        df = df.rename(columns=lambda c: c if c == "Team" else f"{c}{col_suffix}")
        dfs[suffix] = df
        print(f"  ✅ Loaded in-memory data for {suffix} ({df.shape[0]}×{df.shape[1]})")

    if "standings" not in dfs:
        print(f"  ❌ No standings table for {name}; cannot merge.")
        return

    merged = dfs["standings"]
    for key in ["home_away", "standard_stats", "goalkeeping", "shooting", "playing_time", "misc_stats"]:
        if key in dfs:
            merged = pd.merge(merged, dfs[key], on="Team", how="inner")

    print(f"\n  Combined shape: {merged.shape[0]} rows × {merged.shape[1]} cols")
    merged.to_excel(output_path, index=False, sheet_name="Features")
    print(f"  ✅ Saved combined Excel → {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
# Main – iterate over every configured league
# ─────────────────────────────────────────────────────────────────────────────
for league in LEAGUES:
    scrape_league(league)

print("\n✅ All leagues processed successfully!")