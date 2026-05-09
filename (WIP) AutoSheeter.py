import os
import json
import time
import urllib.request
import urllib.error
import configparser
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials


#CONFIG in config.ini


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
config = configparser.ConfigParser()
config_path = os.path.join(BASE_DIR, "config.ini")

if not os.path.exists(config_path):
    raise FileNotFoundError(f"config.ini not found at {config_path}")

config.read(config_path)


json_file_name       = config.get("GoogleSheets", "json_file_name")
SPREADSHEET_ID       = config.get("GoogleSheets", "spreadsheet_id")
SOURCE_SHEET         = config.get("GoogleSheets", "source_sheet")
cell_to_autochange   = config.get("GoogleSheets", "cell_to_autochange")


START_ROW = config.getint("Sheet", "start_row")


COLUMN_MAP = {
    "MoonInfo_Name":        config.get("Columns", "MoonInfo_Name"),
    "MoonInfo_Weather":     config.get("Columns", "MoonInfo_Weather"),
    "DungeonInfo_Interior": config.get("Columns", "DungeonInfo_Interior"),
    "DungeonInfo_ItemCount":config.get("Columns", "DungeonInfo_ItemCount"),
    "CollectedTotal":       config.get("Columns", "CollectedTotal"),
    "BottomLine":           config.get("Columns", "BottomLine"),
    "ValueSold":            config.get("Columns", "ValueSold"),
    "NewQuota":             config.get("Columns", "NewQuota"),
    "ExtraNumber":          config.get("Columns", "ExtraNumber"),
    "Seed":                 config.get("Columns", "Seed"),
}


SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "extra", json_file_name)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build("sheets", "v4", credentials=creds)


result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=f"{SOURCE_SHEET}!{cell_to_autochange}"
).execute()


values = result.get("values")
if not values or not values[0]:
    raise ValueError(f"Cell {cell_to_autochange} is empty or missing")


target_sheet = values[0][0].strip()
print(f"Target sheet from {cell_to_autochange}: '{target_sheet}'")


STATS_URL = os.getenv("STATS_URL", "http://localhost:2145/")
FALLBACK_STATS_FILE = os.path.join(
    os.path.expanduser("~"),
    "Documents",
    "LethalCompanyStats",
    "stats.json"
)


def parse_sse_payload(raw_text):
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    data_lines = []
    for line in lines:
        if line.startswith("data:"):
            data_lines.append(line[len("data:"):].strip())
    return "\n".join(data_lines)


def get_stats_from_http():
    try:
        with urllib.request.urlopen(STATS_URL, timeout=5) as response:
            raw = response.read().decode("utf-8")

        raw = raw.lstrip("\ufeff")
        if not raw.strip():
            print(f"✗ Empty response from {STATS_URL}")
            return None

        try:
            stats_data = json.loads(raw)
        except json.JSONDecodeError:
            payload = parse_sse_payload(raw)
            if not payload.strip():
                print(f"✗ No SSE data payload found in response from {STATS_URL}: {repr(raw[:200])}")
                return None
            stats_data = json.loads(payload)

        print(f"✓ Received stats JSON from {STATS_URL}")
        return stats_data
    except urllib.error.URLError as e:
        print(f"✗ HTTP error fetching stats from {STATS_URL}: {e}")
    except json.JSONDecodeError as e:
        print(f"✗ Error parsing JSON from {STATS_URL}: {e}")
        print(f"  raw response snippet: {repr(raw[:200])}")
    except Exception as e:
        print(f"✗ Unexpected error fetching stats from {STATS_URL}: {e}")
    return None


def get_stats_from_file():
    if not os.path.exists(FALLBACK_STATS_FILE):
        return None
    try:
        with open(FALLBACK_STATS_FILE, "r", encoding="utf-8") as f:
            stats_data = json.load(f)
        print(f"✓ Received stats JSON from local file {FALLBACK_STATS_FILE}")
        return stats_data
    except json.JSONDecodeError as e:
        print(f"✗ Error parsing JSON from local file: {e}")
    except Exception as e:
        print(f"✗ Error reading local stats file: {e}")
    return None


def get_stats():
    stats = get_stats_from_http()
    if stats is not None:
        return stats
    return get_stats_from_file()


def strip_apostrophe(value):
    return str(value).lstrip("'")


def normalize_stats(stats):
    dungeon = stats.get("DungeonInfo") or {}
    moon    = stats.get("MoonInfo") or {}
    BeeInfo_Values    = (stats.get("BeeInfo") or {}).get("Values") or []
    BirdInfo_EggValues = (stats.get("BirdInfo") or {}).get("EggValues") or []
    extra_number = len(BeeInfo_Values) + len(BirdInfo_EggValues)
    return {
        "MoonInfo_Name":         strip_apostrophe(moon.get("Name", "")),
        "MoonInfo_Weather":      strip_apostrophe(moon.get("Weather", "")),
        "DungeonInfo_Interior":  strip_apostrophe(dungeon.get("Interior", "")),
        "DungeonInfo_ItemCount": int(strip_apostrophe(dungeon.get("ItemCount", 0))),
        "CollectedTotal":        int(strip_apostrophe(stats.get("CollectedTotal", 0))),
        "BottomLine":            int(strip_apostrophe(stats.get("BottomLine", 0))),
        "ValueSold":             int(strip_apostrophe(stats.get("ValueSold", 0))),
        "NewQuota":              int(strip_apostrophe(stats.get("NewQuota", 0))),
        "ExtraNumber":           extra_number,
        "Seed":                  strip_apostrophe(stats.get("Seed", ""))
    }


def get_next_empty_row():
    col = COLUMN_MAP["MoonInfo_Name"]
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{target_sheet}!{col}{START_ROW}:{col}1000"
    ).execute()
    rows = result.get("values", [])
    next_row = START_ROW + len(rows)
    print(f"✓ Next empty row determined: {next_row}")
    return next_row


def write_to_cell(value, cell):
    try:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{target_sheet}!{cell}",
            valueInputOption="RAW",
            body={"values": [[value]]}
        ).execute()
        print(f"✓ Wrote '{value}' to {target_sheet}!{cell}")
    except Exception as e:
        print(f"✗ Error writing to {target_sheet}!{cell}: {e}")


def update_sheet_from_stats(stats):
    normalized = normalize_stats(stats)
    target_row = get_next_empty_row()

    if normalized["MoonInfo_Name"] == "71 Gordion":
        print("✓ On 71 Gordion — skipping entirely")
        if normalized["ValueSold"] == 0 or normalized["NewQuota"] == 0:
            print("✓ On 71 Gordion — ValueSold or NewQuota is 0, skipping entirely")
            return
        if normalized["ValueSold"] != 0:
            write_to_cell(normalized["ValueSold"], f'{COLUMN_MAP["ValueSold"]}{target_row - 3}')
        if normalized["NewQuota"] != 0:
            write_to_cell(normalized["NewQuota"], f'{COLUMN_MAP["NewQuota"]}{target_row}')
        return

    for key, col in COLUMN_MAP.items():
        if key in ("ValueSold", "NewQuota") and normalized[key] == 0:
            print(f"✓ {key} is 0, skipping")
            continue
        write_to_cell(normalized[key], f"{col}{target_row}")


def main():
    last_stats_text = None
    while True:
        stats = get_stats()
        if stats is not None:
            current_stats_text = json.dumps(stats, sort_keys=True)
            if current_stats_text != last_stats_text:
                update_sheet_from_stats(stats)
                last_stats_text = current_stats_text
            else:
                print("✓ No change in stats, skipping update")
        time.sleep(1)


if __name__ == "__main__":
    main()