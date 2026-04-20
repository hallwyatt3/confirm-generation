import requests
import sqlite3
import time
from pathlib import Path
from datetime import datetime

from generate_confirm import generate_confirm

SMARTSHEET_ID = "934997499268996"
API_TOKEN     = "96Jh9H2na9PfQCY1ruVKK0vos5RZwaGQBLMqZ"
DB_PATH       = Path(__file__).parent / "trade_book.db"
POLL_SECONDS  = 30 * 60  # every 30 minutes

headers = {"Authorization": f"Bearer {API_TOKEN}"}


def fetch_rows():
    resp = requests.get(
        f"https://api.smartsheet.com/2.0/sheets/{SMARTSHEET_ID}", headers=headers
    )
    resp.raise_for_status()
    data    = resp.json()
    columns = {col["id"]: col["title"] for col in data.get("columns", [])}
    flat_rows = []
    for row in data.get("rows", []):
        flat = {"smartsheet_row_id": str(row["id"])}
        for cell in row.get("cells", []):
            col_name = columns.get(cell.get("columnId"), "")
            if col_name:
                flat[col_name] = cell.get("displayValue") or cell.get("value")
        flat_rows.append(flat)
    return flat_rows, list(columns.values())


def init_db(con, col_names):
    col_defs = ", ".join(f'"{c}" TEXT' for c in col_names)
    con.execute(f"""
        CREATE TABLE IF NOT EXISTS trades (
            smartsheet_row_id TEXT UNIQUE,
            {col_defs},
            synced_at TEXT DEFAULT (datetime('now'))
        )
    """)
    con.commit()


def get_known_ids(con):
    rows = con.execute("SELECT smartsheet_row_id FROM trades").fetchall()
    return {r[0] for r in rows}


def save_rows(con, flat_rows, col_names):
    placeholders = ", ".join("?" * (len(col_names) + 1))
    col_list     = "smartsheet_row_id, " + ", ".join(f'"{c}"' for c in col_names)
    for flat in flat_rows:
        values = [flat.get("smartsheet_row_id")] + [flat.get(c) for c in col_names]
        con.execute(
            f"INSERT OR REPLACE INTO trades ({col_list}) VALUES ({placeholders})",
            values,
        )
    con.commit()


def sync():
    print(f"[{datetime.now():%H:%M:%S}] Checking Smartsheet...")

    flat_rows, col_names = fetch_rows()
    con = sqlite3.connect(DB_PATH)
    init_db(con, col_names)

    known_ids = get_known_ids(con)
    new_rows  = [r for r in flat_rows if r["smartsheet_row_id"] not in known_ids]

    if not new_rows:
        print(f"  No new rows. ({len(flat_rows)} total)")
        con.close()
        return

    print(f"  {len(new_rows)} new row(s):")
    for row in new_rows:
        print(
            f"    Trade {row.get('Trade ID')} | {row.get('Counterparty')} | "
            f"{row.get('Direction')} | {row.get('Volume (MMBtu/d)')} MMBtu/d"
        )
        generate_confirm(row)

    save_rows(con, new_rows, col_names)
    con.close()


if __name__ == "__main__":
    print("Smartsheet monitor started (interval: 30 min).")
    print("PDFs will be saved to: confirmations/")
    try:
        sync()
    except Exception as e:
        print(f"  [ERROR] {e}")
    time.sleep(POLL_SECONDS)
