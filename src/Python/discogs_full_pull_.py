"""
Scrape personal collection data from Discogs and save as an Excel table with lowest price
"""

import requests
import time
import sys
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# -- config --
API_TOKEN = "YOUR_API_KEY"
USERNAME = "soupbadger"
OUTPUT = "collection_roi_table.xlsx"
REQ_SLEEP = 1.1
PER_PAGE = 100
AUTOSAVE_INTERVAL = 10  # autosave after every N releases processed
# --------------

COLLECTION_BASE = f"https://api.discogs.com/users/{USERNAME}/collection/folders/0/releases"
MARKET_STATS = "https://api.discogs.com/marketplace/stats/"

headers = {
    "Authorization": f"Discogs token={API_TOKEN}",
    "User-Agent": "DiscogsROIScript/1.0"
}

def get_lowest(stats_json):
    if not stats_json:
        return None
    lp = stats_json.get("lowest_price")
    if lp is None:
        return None
    if isinstance(lp, dict):
        try:
            return float(lp.get("value"))
        except Exception:
            return None
    try:
        return float(lp)
    except Exception:
        return None

def json_retries(url, max_retries=5):
    backoff = 1.0
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=30)
        except requests.RequestException:
            if attempt == max_retries:
                return None
            time.sleep(backoff)
            backoff *= 2
            continue

        if resp.status_code == 200:
            try:
                return resp.json()
            except Exception:
                return None

        if resp.status_code == 429:
            ra = resp.headers.get("Retry-After")
            try:
                wait = float(ra) if ra else backoff
            except Exception:
                wait = backoff
            time.sleep(wait)
            backoff *= 2
            continue

        if 500 <= resp.status_code < 600:
            time.sleep(backoff)
            backoff *= 2
            continue

        return None
    return None

def autosave_results_table(with_price, output_path):
    """Save results as an Excel table using openpyxl"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Collection"

    headers = ["release_id", "artist", "title", "lowest_price", "num_for_sale"]
    ws.append(headers)

    for row in with_price:
        ws.append([
            row.get("release_id"),
            row.get("artist"),
            row.get("title"),
            row.get("lowest_price"),
            row.get("num_for_sale")
        ])

    # create Excel table
    tab = Table(displayName="CollectionTable", ref=f"A1:E{len(with_price)+1}")
    style = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=True, showColumnStripes=False
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)

    # auto column width
    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    wb.save(output_path)
    print(f"Saved {len(with_price)} rows to {output_path}")

def main():
    results_with = []
    url = f"{COLLECTION_BASE}?per_page={PER_PAGE}&page=1"
    total_seen = 0

    while url:
        col_json = json_retries(url)
        if not col_json:
            print(f"Failed to fetch collection page: {url}", file=sys.stderr)
            break

        releases = col_json.get("releases", [])
        for r in releases:
            total_seen += 1
            basic = r.get("basic_information", {})
            release_id = basic.get("id")
            title = basic.get("title", "")
            artists = basic.get("artists", [])
            artist_names = ", ".join(a.get("name", "") for a in artists) if artists else ""

            if not release_id:
                print(f"[{total_seen}] missing id -> skipping")
            else:
                stats_url = MARKET_STATS + str(release_id)
                stats_json = json_retries(stats_url)
                lowest = get_lowest(stats_json)

                if lowest is not None:
                    lowest_2 = round(lowest, 2)
                    num_for_sale = stats_json.get("num_for_sale") if stats_json else None

                    results_with.append({
                        "release_id": release_id,
                        "artist": artist_names,
                        "title": title,
                        "lowest_price": lowest_2,
                        "num_for_sale": num_for_sale
                    })
                    print(f"[{total_seen}] {release_id} | {artist_names} - {title} => L:{lowest_2}")

            if total_seen % AUTOSAVE_INTERVAL == 0:
                autosave_results_table(results_with, OUTPUT)

            time.sleep(REQ_SLEEP)

        # pagination
        pagination = col_json.get("pagination", {})
        urls = pagination.get("urls", {}) if pagination else {}
        next_url = urls.get("next")
        if not next_url:
            break
        url = next_url
        time.sleep(REQ_SLEEP)

    autosave_results_table(results_with, OUTPUT)

if __name__ == "__main__":
    main()
