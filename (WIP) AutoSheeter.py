import os
import json
import time
import urllib.request
import urllib.error
import configparser
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
config = configparser.ConfigParser()
config_path = os.path.join(BASE_DIR, "config.ini")

if not os.path.exists(config_path):
    raise FileNotFoundError(f"config.ini not found at {config_path}")

config.read(config_path)

json_file_name     = config.get("GoogleSheets", "json_file_name")
SPREADSHEET_ID     = config.get("GoogleSheets", "spreadsheet_id")
SOURCE_SHEET       = config.get("GoogleSheets", "source_sheet")
cell_to_autochange = config.get("GoogleSheets", "cell_to_autochange")

START_ROW = config.getint("Sheet", "start_row")

PLAYER_COLUMNS = [c.strip() for c in config.get("Columns", "Players").split(",") if c.strip()]

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
    "SIDType":              config.get("Columns", "SID"),
    "InfestationType":      config.get("Columns", "Infestation"),
    "IndoorFog":            config.get("Columns", "IndoorFog"),
    "MeteorShower":         config.get("Columns", "MeteorShower"),
    "BeehiveAmount":        config.get("Columns", "BeehiveAmount"),
    "BeehiveValue":         config.get("Columns", "BeehiveValue"),
    "EggAmount":            config.get("Columns", "EggAmount", fallback=None),
    "EggValue":             config.get("Columns", "EggValue"),
}

CHECKBOX_FIELDS = {"SIDType", "InfestationType", "IndoorFog", "MeteorShower", "EggAmount"}

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
print(f"Target sheet: '{target_sheet}'")

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
            return None
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            payload = parse_sse_payload(raw)
            if not payload.strip():
                return None
            return json.loads(payload)
    except Exception:
        return None


def get_stats_from_file():
    if not os.path.exists(FALLBACK_STATS_FILE):
        return None
    try:
        with open(FALLBACK_STATS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def get_stats():
    stats = get_stats_from_http()
    if stats is not None:
        return stats
    return get_stats_from_file()


def strip_apostrophe(value):
    return str(value).lstrip("'")


def strip_moon_number(name: str) -> str:
    parts = name.split(" ", 1)
    if len(parts) == 2 and parts[0].rstrip("-").isdigit():
        return parts[1]
    return name


def normalize_players(raw_players: dict) -> list[dict]:
    players = []
    for steam_id, data in raw_players.items():
        name           = strip_apostrophe(data.get("Name", steam_id))
        alive          = data.get("Alive", False)
        disconnected   = data.get("Disconnected", False)
        time_of_death  = strip_apostrophe(data.get("TimeOfDeath", "")).strip()
        cause_of_death = strip_apostrophe(data.get("CauseOfDeath", "")).strip()

        if disconnected:
            status = "DC"
        elif cause_of_death == "Abandonment":
            status = "M"
        elif alive:
            status = "S"
        else:
            status = "X"

        note_parts = [name]
        if time_of_death:
            note_parts.append(f"Time of Death: {time_of_death}")
        if cause_of_death:
            note_parts.append(f"Cause of Death: {cause_of_death}")
        note = "\n".join(note_parts)

        players.append({"status": status, "note": note})

    return players


def normalize_stats(stats):
    dungeon   = stats.get("DungeonInfo") or {}
    moon      = stats.get("MoonInfo") or {}
    bee_info  = stats.get("BeeInfo") or {}
    bird_info = stats.get("BirdInfo") or {}

    bee_values = bee_info.get("Values") or []
    egg_values = bird_info.get("EggValues") or []

    raw_players = stats.get("Players") or {}
    if not isinstance(raw_players, dict):
        raw_players = {}

    moon_name = strip_moon_number(strip_apostrophe(moon.get("Name", "")))
    weather   = strip_apostrophe(moon.get("Weather", ""))
    if weather == "Mild":
        weather = "Clear"

    indoor_fog_val = "true" if stats.get("IndoorFog", False) else ""

    meteor_time = strip_apostrophe(stats.get("MeteorShowerTime", "")).strip()
    meteor_val  = meteor_time if meteor_time else ""

    bee_min = sorted([int(v) for v in bee_values if int(v) < 100])
    bee_max = sorted([int(v) for v in bee_values if int(v) >= 100])
    beehive_amount = f"{len(bee_min)}|{len(bee_max)}" if bee_values else ""
    beehive_value  = f"{bee_min[0] if bee_min else 0}|{bee_max[0] if bee_max else 0}" if bee_values else ""

    egg_amount_val  = "true" if egg_values else ""
    egg_value_total = sum(int(v) for v in egg_values) if egg_values else 0

    return {
        "MoonInfo_Name":         moon_name,
        "MoonInfo_Weather":      weather,
        "DungeonInfo_Interior":  strip_apostrophe(dungeon.get("Interior", "")),
        "DungeonInfo_ItemCount": int(strip_apostrophe(dungeon.get("ItemCount", 0))),
        "CollectedTotal":        int(strip_apostrophe(stats.get("CollectedTotal", 0))),
        "BottomLine":            int(strip_apostrophe(stats.get("BottomLine", 0))),
        "ValueSold":             int(strip_apostrophe(stats.get("ValueSold", 0))),
        "NewQuota":              int(strip_apostrophe(stats.get("NewQuota", 0))),
        "ExtraNumber":           len(bee_values) + len(egg_values),
        "Seed":                  strip_apostrophe(stats.get("Seed", "")),
        "SIDType":               strip_apostrophe(stats.get("SIDType", "")),
        "InfestationType":       strip_apostrophe(stats.get("InfestationType", "")),
        "IndoorFog":             indoor_fog_val,
        "MeteorShower":          meteor_val,
        "BeehiveAmount":          beehive_amount,
        "BeehiveValue":          beehive_value,
        "EggAmount":             egg_amount_val,
        "EggValue":              egg_value_total,
        "Players":               normalize_players(raw_players),
    }


def col_letter_to_index(col: str) -> int:
    col = col.upper()
    index = 0
    for ch in col:
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def get_sheet_id(sheet_name: str) -> int:
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props["sheetId"]
    raise ValueError(f"Sheet '{sheet_name}' not found in spreadsheet")


def get_next_empty_row():
    col = COLUMN_MAP["MoonInfo_Name"]
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{target_sheet}!{col}{START_ROW}:{col}1000"
    ).execute()
    rows = result.get("values", [])
    return START_ROW + len(rows)


def write_to_cell(value, cell):
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{target_sheet}!{cell}",
        valueInputOption="RAW",
        body={"values": [[value]]}
    ).execute()


def write_cell_with_note(col: str, row: int, value: str, note: str):
    sheet_id  = get_sheet_id(target_sheet)
    col_index = col_letter_to_index(col)
    row_index = row - 1

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "updateCells": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    row_index,
                    "endRowIndex":      row_index + 1,
                    "startColumnIndex": col_index,
                    "endColumnIndex":   col_index + 1,
                },
                "rows": [{"values": [{"userEnteredValue": {"stringValue": value}, "note": note}]}],
                "fields": "userEnteredValue,note",
            }
        }]}
    ).execute()


def write_checkbox_with_note(col: str, row: int, note_text: str):
    sheet_id  = get_sheet_id(target_sheet)
    col_index = col_letter_to_index(col)
    row_index = row - 1
    checked   = bool(note_text.strip())

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "updateCells": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    row_index,
                    "endRowIndex":      row_index + 1,
                    "startColumnIndex": col_index,
                    "endColumnIndex":   col_index + 1,
                },
                "rows": [{"values": [{"userEnteredValue": {"boolValue": checked}, "note": note_text if checked else ""}]}],
                "fields": "userEnteredValue,note",
            }
        }]}
    ).execute()


def write_players(players: list[dict], row: int):
    if len(players) > len(PLAYER_COLUMNS):
        print(f"⚠ More players ({len(players)}) than configured columns ({len(PLAYER_COLUMNS)}); extras ignored")
    for i, player in enumerate(players):
        if i >= len(PLAYER_COLUMNS):
            break
        write_cell_with_note(PLAYER_COLUMNS[i], row, player["status"], player["note"])


def update_sheet_from_stats(stats):
    normalized = normalize_stats(stats)
    target_row = get_next_empty_row()

    if normalized["MoonInfo_Name"] == "71 Gordion":
        if normalized["ValueSold"] == 0 or normalized["NewQuota"] == 0:
            return
        if normalized["ValueSold"] != 0:
            write_to_cell(normalized["ValueSold"], f'{COLUMN_MAP["ValueSold"]}{target_row - 3}')
        if normalized["NewQuota"] != 0:
            write_to_cell(normalized["NewQuota"], f'{COLUMN_MAP["NewQuota"]}{target_row}')
        print(f"Updated {target_sheet} (Gordion: sold/quota)")
        return

    for key, col in COLUMN_MAP.items():
        if col is None:
            continue
        value = normalized[key]

        if key in CHECKBOX_FIELDS:
            write_checkbox_with_note(col, target_row, str(value))
            continue

        if key in ("ValueSold", "NewQuota") and value == 0:
            continue

        if key == "EggValue" and value == 0:
            write_to_cell("X", f"{col}{target_row}")
            continue

        if key in ("BeehiveAmount", "BeehiveValue") and value == "":
            write_to_cell("X", f"{col}{target_row}")
            continue

        write_to_cell(value, f"{col}{target_row}")

    write_players(normalized["Players"], target_row)
    print(f"Updated {target_sheet} (row {target_row})")


def main():
    print(f"Watching for stats — target sheet: '{target_sheet}'")
    last_stats_text = None
    while True:
        try:
            stats = get_stats()
            if stats is not None:
                current_stats_text = json.dumps(stats, sort_keys=True)
                if current_stats_text != last_stats_text:
                    update_sheet_from_stats(stats)
                    last_stats_text = current_stats_text
        except Exception as e:
            print(f"✗ Error: {e}")
        time.sleep(1)


if __name__ == "__main__":
    main()