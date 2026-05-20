import os
import re
import json
import time
import urllib.request
import configparser
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
config = configparser.ConfigParser()
config_path = os.path.join(BASE_DIR, "config.ini")

if not os.path.exists(config_path):
    raise FileNotFoundError(f"config.ini not found at {config_path}")

config.read(config_path)


def cfg_get(section, key, fallback=None):
    try:
        value = config.get(section, key, fallback=fallback)
    except (configparser.NoSectionError, configparser.NoOptionError):
        return None
    if value is None or str(value).strip().lower() in ("", "none"):
        return None
    return str(value).strip()


def cfg_get_int(section, key, fallback=None):
    raw = cfg_get(section, key)
    if raw is None:
        return fallback
    try:
        return int(raw)
    except ValueError:
        return fallback


json_file_name = cfg_get("GoogleSheets", "json_file_name")
SPREADSHEET_ID = cfg_get("GoogleSheets", "spreadsheet_id")
target_sheet   = cfg_get("GoogleSheets", "target_sheet")
START_ROW      = cfg_get_int("Sheet", "start_row", fallback=2)

if not json_file_name:
    raise ValueError("config.ini: [GoogleSheets] json_file_name is required.")
if not SPREADSHEET_ID:
    raise ValueError("config.ini: [GoogleSheets] spreadsheet_id is required.")
if not target_sheet:
    raise ValueError("config.ini: [GoogleSheets] target_sheet is required.")

raw_players_cfg = cfg_get("Columns", "Players")
PLAYER_COLUMNS  = [c.strip() for c in raw_players_cfg.split(",") if c.strip()] if raw_players_cfg else []

COLUMN_MAP = {
    "NewQuota":              cfg_get("Columns", "NewQuota"),
    "MoonInfo_Name":         cfg_get("Columns", "MoonInfo_Name"),
    "MoonInfo_Weather":      cfg_get("Columns", "MoonInfo_Weather"),
    "DungeonInfo_Interior":  cfg_get("Columns", "DungeonInfo_Interior"),
    "DungeonInfo_ItemCount": cfg_get("Columns", "DungeonInfo_ItemCount"),
    "BeehiveAmount":         cfg_get("Columns", "BeehiveAmount"),
    "BeehiveValue":          cfg_get("Columns", "BeehiveValue"),
    "BeehiveCollected":      cfg_get("Columns", "BeehiveCollected"),
    "EggValue":              cfg_get("Columns", "EggValue"),
    "KnifeInfo":             cfg_get("Columns", "KnifeInfo"),
    "ShotgunInfo":           cfg_get("Columns", "ShotgunInfo"),
    "CollectedTotal":        cfg_get("Columns", "CollectedTotal"),
    "BottomLine":            cfg_get("Columns", "BottomLine"),
    "Scan":                  cfg_get("Columns", "Scan"),
    "OutsideItemsValue":     cfg_get("Columns", "OutsideItemsValue"),
    "ValueSold":             cfg_get("Columns", "ValueSold"),
    "SIDType":               cfg_get("Columns", "SID"),
    "InfestationType":       cfg_get("Columns", "Infestation"),
    "AppyLess":              cfg_get("Columns", "AppyLess"),
    "IndoorFog":             cfg_get("Columns", "IndoorFog"),
    "MeteorShower":          cfg_get("Columns", "MeteorShower"),
    "GiftBoxes":             cfg_get("Columns", "GiftBoxes"),
    "Seed":                  cfg_get("Columns", "Seed"),
}

SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "extra", json_file_name)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds   = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build("sheets", "v4", credentials=creds)

_sheet_id_cache = {}

STATS_URL           = os.getenv("STATS_URL", "http://localhost:2145/")
FALLBACK_STATS_FILE = os.path.join(os.path.expanduser("~"), "Documents", "LethalCompanyStats", "stats.json")

print(f"Target sheet: '{target_sheet}'")
disabled = [k for k, v in COLUMN_MAP.items() if v is None]
if disabled:
    print(f"Columns disabled: {', '.join(disabled)}")
if not PLAYER_COLUMNS:
    print("Player columns disabled.")


def col_letter_to_index(col):
    col = col.upper()
    index = 0
    for ch in col:
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def sorted_column_map_keys(column_map):
    return sorted(column_map, key=lambda k: col_letter_to_index(column_map[k]) if column_map[k] else 10**9)


def get_sheet_id(sheet_name):
    if sheet_name not in _sheet_id_cache:
        meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        for sheet in meta.get("sheets", []):
            props = sheet.get("properties", {})
            _sheet_id_cache[props["title"]] = props["sheetId"]
    if sheet_name not in _sheet_id_cache:
        raise ValueError(f"Sheet '{sheet_name}' not found in spreadsheet")
    return _sheet_id_cache[sheet_name]


def parse_sse_payload(raw_text):
    data_lines = [l[5:].strip() for l in raw_text.splitlines() if l.strip().startswith("data:")]
    return "\n".join(data_lines)


def get_stats():
    try:
        with urllib.request.urlopen(STATS_URL, timeout=5) as r:
            raw = r.read().decode("utf-8").lstrip("\ufeff")
        if raw.strip():
            try:
                return json.loads(raw)
            except json.JSONDecodeError:
                payload = parse_sse_payload(raw)
                if payload.strip():
                    return json.loads(payload)
    except Exception:
        pass
    if os.path.exists(FALLBACK_STATS_FILE):
        try:
            with open(FALLBACK_STATS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return None


def strip_apostrophe(value):
    return str(value).lstrip("'")


def coerce_value(value):
    if isinstance(value, bool):
        return value
    for cast in (int, float):
        try:
            return cast(value)
        except (ValueError, TypeError):
            pass
    return value


def make_cell_value(value):
    if isinstance(value, bool):
        return {"boolValue": value}
    for cast in (int, float):
        try:
            return {"numberValue": cast(value)}
        except (ValueError, TypeError):
            pass
    return {"stringValue": str(value)}


def build_update_request(sheet_id, col, row, user_entered_value, note=None):
    ci = col_letter_to_index(col)
    ri = row - 1
    row_val = {"userEnteredValue": user_entered_value}
    fields  = "userEnteredValue"
    if note is not None:
        row_val["note"] = note
        fields += ",note"
    return {"updateCells": {
        "range":  {"sheetId": sheet_id, "startRowIndex": ri, "endRowIndex": ri + 1,
                   "startColumnIndex": ci, "endColumnIndex": ci + 1},
        "rows":   [{"values": [row_val]}],
        "fields": fields,
    }}


def normalize_players(raw_players):
    players = []
    names = []
    for steam_id, data in sorted(raw_players.items(), key=lambda item: int(item[0])):
        cause_of_death = strip_apostrophe(data.get("CauseOfDeath", "")).strip()
        time_of_death  = strip_apostrophe(data.get("TimeOfDeath", "")).strip()
        names.append(data.get("Name", steam_id))

        if data.get("Disconnected"):
            status = "DC"
        elif cause_of_death.lower() in ("abandonment", "abandoned"):
            status = "M"
        elif data.get("Alive"):
            status = "S"
        else:
            late_death = False
            if time_of_death:
                try:
                    h, m = map(int, time_of_death.split(":")[:2])
                    late_death = (h == 22 and m >= 45) or h >= 23
                except (ValueError, AttributeError):
                    pass
            status = "SX" if late_death else "X"

        note = ""
        if status != "M":
            parts = []
            if time_of_death:
                parts.append(f"Time of Death: {time_of_death}")
            if cause_of_death:
                parts.append(f"Cause of Death: {cause_of_death}")
            note = "\n".join(parts)

        players.append({"status": status, "note": note})
    return players, names


def normalize_gift_boxes(raw):
    if not raw:
        return {"collected_any": False, "cell_value": "", "note": ""}
    collected = [b for b in raw if b.get("Collected")]
    missed    = [b for b in raw if not b.get("Collected")]
    if collected:
        net        = sum(int(b.get("NewScrapValue", 0)) - int(b.get("GiftScrapValue", 0)) for b in collected)
        cell_value = f"+{net}" if net >= 0 else str(net)
    else:
        cell_value = ""
    note = "\n".join(
        f"Gift {i}: Box: {int(b.get('GiftScrapValue', 0))} ; Item: {int(b.get('NewScrapValue', 0))}"
        for i, b in enumerate(missed, 1)    
    )
    return {"collected_any": bool(collected), "cell_value": cell_value, "note": note}   


def normalize_missed_items(raw):
    if not raw:
        return {"cell_value": "", "note": ""}
    uncollected = [i for i in raw if not i.get("CollectedOnPreviousDay")]
    if not uncollected:
        return {"cell_value": "", "note": ""}
    note = "\n".join(f"{i.get('ItemType', 'Unknown')}: {int(i.get('Value', 0))}" for i in uncollected)
    return {"cell_value": str(len(uncollected)), "note": note}



def normalize_weapon_count(raw_info, label):
    if not raw_info:
        return {"cell_value": "0", "note": ""}
    collected  = raw_info.get("Collected") or []
    available  = raw_info.get("Available") or []
    missed     = available[len(collected):]
    note       = " ; ".join(f"{label}: {int(v)}" for v in missed) if missed else ""
    return {"cell_value": str(len(collected)), "note": note}


def normalize_beehive_collected(bee_info, new_bee_format):
    collected = [int(v) for v in (bee_info.get("Collected") or [])]
    if not collected:
        return ""
    if new_bee_format:
        return f"{sum(1 for v in collected if v < 100)}|{sum(1 for v in collected if v >= 100)}"
    return str(len(collected))


def normalize_outside_items_value(bee_info, egg_info, new_bee_format):
    bee_avail = [int(v) for v in (bee_info.get("Available") or [])]
    bee_coll  = [int(v) for v in (bee_info.get("Collected") or [])]
    egg_avail = [int(v) for v in (egg_info.get("Available") or [])]
    egg_coll  = [int(v) for v in (egg_info.get("Collected") or [])]

    total            = sum(bee_coll) + sum(egg_coll)
    bee_missed_small = sum(1 for v in bee_avail if v < 100)  - sum(1 for v in bee_coll if v < 100)
    bee_missed_large = sum(1 for v in bee_avail if v >= 100) - sum(1 for v in bee_coll if v >= 100)
    bee_missed_total = sum(1 for v in bee_avail) - sum(1 for v in bee_coll)

    remaining_eggs = sorted(egg_avail)
    for v in sorted(egg_coll):
        if v in remaining_eggs:
            remaining_eggs.remove(v)

    note_parts = []
    if new_bee_format:
        if bee_missed_small > 0 or bee_missed_large > 0:
            note_parts.append(f"Bee ({bee_missed_small}|{bee_missed_large})")
    else:
        if bee_missed_total > 0:
            note_parts.append(f"Bee ({bee_missed_total})")
    if remaining_eggs:
        note_parts.append(f"Egg ({', '.join(str(v) for v in remaining_eggs)})")

    return {
        "cell_value": str(total) if total > 0 else "X",
        "note": ("Missing: " + " ".join(note_parts)) if note_parts else "",
    }


def normalize_stats(stats):
    dungeon      = stats.get("DungeonInfo") or {}
    moon         = stats.get("MoonInfo") or {}
    bee_info     = stats.get("BeeInfo") or {}
    egg_info     = stats.get("EggInfo") or {}
    missed_items = stats.get("MissedItems") or []

    bee_avail = [int(v) for v in (bee_info.get("Available") or [])]
    egg_avail = [int(v) for v in (egg_info.get("Available") or [])]

    weather = strip_apostrophe(moon.get("Weather", ""))
    if weather == "Mild":
        weather = "Clear"

    moon_name = strip_apostrophe(moon.get("Name", ""))
    parts = moon_name.split(" ", 1)
    if len(parts) == 2 and parts[0].rstrip("-").isdigit():
        moon_name = parts[1]

    interior = strip_apostrophe(dungeon.get("Interior", ""))
    interior = re.sub(r'flow', '', interior, flags=re.IGNORECASE)
    interior = re.sub(r'([a-z])([A-Z])', r'\1 \2', interior)
    interior = re.sub(r'\d+', '', interior)
    interior = re.sub(r' {2,}', ' ', interior).strip()

    version        = int(stats.get("Version", 0))
    new_bee_format = version >= 70

    if bee_avail:
        if new_bee_format:
            bee_small      = [v for v in bee_avail if v < 100]
            bee_large      = [v for v in bee_avail if v >= 100]
            beehive_amount = f"{len(bee_small)}|{len(bee_large)}"
            beehive_value  = f"{bee_small[0] if bee_small else 0}|{bee_large[0] if bee_large else 0}"
        else:
            beehive_amount = str(len(bee_avail))
            beehive_value  = str(bee_avail[0])
    else:
        beehive_amount = ""
        beehive_value  = ""

    raw_players = stats.get("Players") or {}
    if not isinstance(raw_players, dict):
        raw_players = {}
    players, player_names = normalize_players(raw_players)

    return {
        "NewQuota":              int(strip_apostrophe(stats.get("NewQuota", 0))),
        "MoonInfo_Name":         moon_name,
        "MoonInfo_Weather":      weather,
        "DungeonInfo_Interior":  interior,
        "DungeonInfo_ItemCount": int(strip_apostrophe(dungeon.get("ItemCount", 0))),
        "BeehiveAmount":         beehive_amount,
        "BeehiveValue":          beehive_value,
        "BeehiveCollected":      normalize_beehive_collected(bee_info, new_bee_format),
        "EggValue":              "|".join(str(v) for v in sorted(egg_avail)) if egg_avail else "",
        "KnifeInfo":             normalize_weapon_count(stats.get("KnifeInfo"), "Knife"),
        "ShotgunInfo":           normalize_weapon_count(stats.get("ShotgunInfo"), "Shotgun"),
        "CollectedTotal":        int(strip_apostrophe(stats.get("CollectedTotal", 0))),
        "BottomLine":            int(strip_apostrophe(stats.get("BottomLine", 0))),
        "Scan":                  normalize_missed_items(missed_items),
        "OutsideItemsValue":     normalize_outside_items_value(bee_info, egg_info, new_bee_format),
        "ValueSold":             int(strip_apostrophe(stats.get("ValueSold", 0))),
        "SIDType":               strip_apostrophe(stats.get("SIDType", "")),
        "InfestationType":       strip_apostrophe(stats.get("InfestationType", "")),
        "AppyLess":              stats.get("AppSpawned", False) if interior == "Facility" else None,
        "IndoorFog":             stats.get("IndoorFog", False),
        "MeteorShower":          strip_apostrophe(stats.get("MeteorShowerTime", "")).strip(),
        "GiftBoxes":             normalize_gift_boxes(stats.get("GiftBoxesOpened") or []),
        "Seed":                  strip_apostrophe(stats.get("Seed", "")),
        "Players":               players,
        "PlayerNames":           player_names,
    }


def get_next_empty_row():
    ANCHOR_KEYS = ["MoonInfo_Name", "MoonInfo_Weather", "DungeonInfo_Interior",
                   "DungeonInfo_ItemCount", "ValueSold", "BottomLine"]
    anchor_cols = [COLUMN_MAP[k] for k in ANCHOR_KEYS if COLUMN_MAP.get(k)]
    if not anchor_cols:
        raise ValueError("None of the anchor columns are configured.")

    occupied = set()
    for col in anchor_cols:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{target_sheet}!{col}{START_ROW}:{col}1000",
            majorDimension="COLUMNS",
        ).execute()
        col_values = (result.get("values") or [[]])[0]
        for rel_idx, cell in enumerate(col_values):
            if str(cell).strip():
                occupied.add(START_ROW + rel_idx)

    max_row = max(occupied, default=START_ROW - 1)
    for row in range(START_ROW, max_row + 2):
        if row not in occupied:
            return row


def update_sheet_from_stats(stats):
    normalized = normalize_stats(stats)
    target_row = get_next_empty_row()
    moon_name  = normalized["MoonInfo_Name"]
    sheet_id   = get_sheet_id(target_sheet)
    requests   = []

    def queue(col, uev, note=None):
        requests.append(build_update_request(sheet_id, col, target_row, uev, note))

    if "gordion" in moon_name.lower() or "galetry" in moon_name.lower():
        value_sold = normalized["ValueSold"]
        new_quota  = normalized["NewQuota"]
        if value_sold == 0 and new_quota == 0:
            return
        if value_sold != 0 and COLUMN_MAP.get("ValueSold"):
            requests.append(build_update_request(sheet_id, COLUMN_MAP["ValueSold"], target_row - 3, make_cell_value(value_sold)))
        if new_quota != 0 and COLUMN_MAP.get("NewQuota"):
            requests.append(build_update_request(sheet_id, COLUMN_MAP["NewQuota"], target_row, make_cell_value(new_quota)))
        service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests}).execute()
        print(f"Updated {target_sheet} (Gordion: sold={value_sold}, quota={new_quota})")
        return

    for key in sorted_column_map_keys(COLUMN_MAP):
        col = COLUMN_MAP[key]
        if col is None:
            continue

        if key == "GiftBoxes":
            gift = normalized["GiftBoxes"]
            queue(col, make_cell_value(gift["cell_value"] if gift["collected_any"] else "X"), gift["note"] or None)
            continue

        if key == "Scan":
            missed = normalized["Scan"]
            queue(col, make_cell_value(missed["cell_value"] or "X"), missed["note"] or None)
            continue

        if key == "OutsideItemsValue":
            outside = normalized["OutsideItemsValue"]
            queue(col, make_cell_value(outside["cell_value"]), outside["note"] or None)
            continue

        if key in ("KnifeInfo", "ShotgunInfo"):
            weapon = normalized[key]
            queue(col, make_cell_value(coerce_value(weapon["cell_value"])), weapon["note"] or None)
            continue

        if key == "AppyLess":
            val = normalized["AppyLess"]
            if val is None:
                continue
            queue(col, {"boolValue": not bool(val)})
            continue

        if key == "IndoorFog":
            queue(col, {"boolValue": bool(normalized[key])})
            continue

        if key in ("MeteorShower", "SIDType", "InfestationType"):
            val     = normalized[key]
            checked = bool(str(val).strip())
            queue(col, {"boolValue": checked}, str(val) if checked else None)
            continue

        value = normalized[key]

        if key in ("ValueSold", "NewQuota") and value == 0:
            continue

        if key in ("EggValue", "BeehiveAmount", "BeehiveValue", "BeehiveCollected") and value == "":
            queue(col, make_cell_value("X"))
            continue

        queue(col, make_cell_value(coerce_value(value)))

    if PLAYER_COLUMNS:
        players      = normalized["Players"]
        player_names = normalized["PlayerNames"]
        if len(players) > len(PLAYER_COLUMNS):
            print(f"More players ({len(players)}) than columns ({len(PLAYER_COLUMNS)}); extras ignored")

        sorted_pcols = sorted(PLAYER_COLUMNS, key=col_letter_to_index)
        header_row   = START_ROW - 1

        header_range = f"{target_sheet}!{sorted_pcols[0]}{header_row}:{sorted_pcols[-1]}{header_row}"
        header_result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=header_range
        ).execute()
        existing_names = (header_result.get("values") or [[]])[0]

        name_requests = []
        for i, pcol in enumerate(sorted_pcols):
            if i >= len(players):
                break
            name = player_names[i] if i < len(player_names) else ""
            existing = existing_names[i] if i < len(existing_names) else ""
            if name != existing:
                name_requests.append(build_update_request(sheet_id, pcol, header_row, make_cell_value(""), name or None))

        if name_requests:
            service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": name_requests}).execute()

        for i, pcol in enumerate(sorted_pcols):
            if i >= len(players):
                break
            p = players[i]
            requests.append(build_update_request(sheet_id, pcol, target_row, make_cell_value(p["status"]), p["note"] or None))

    if requests:
        service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests}).execute()
    print(f"Updated {target_sheet} (row {target_row})")


def main():
    print(f"Watching for stats — target sheet: '{target_sheet}'")
    last_stats_text = None
    while True:
        try:
            stats = get_stats()
            if stats is not None:
                current = json.dumps(stats, sort_keys=True)
                if current != last_stats_text:
                    update_sheet_from_stats(stats)
                    last_stats_text = current
        except Exception as e:
            print(f"✗ Error: {e}")
        time.sleep(1)


if __name__ == "__main__":
    main()