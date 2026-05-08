import os
import json
import time
import urllib.request
import urllib.error
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

json_file_name = "" #json file for the googlesheet api.

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "extra", json_file_name)
 
SPREADSHEET_ID = '' #set ur spreadsheet id here (you get it from the url when in the spreadsheet)
SOURCE_SHEET = '' #the tab of your streamoverlay, so that it auto changes to the current tab you are currently in

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)

cell_to_autochange = 'A30' #the cell that it checks to change to the tab you are currently using for streamoverlay

result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=f'{SOURCE_SHEET}!{cell_to_autochange}'
).execute()

values = result.get('values')

if not values or not values[0]:
    raise ValueError(f"Cell {cell_to_autochange} is empty or missing")

target_sheet = values[0][0].strip()

print(f"Target sheet from {cell_to_autochange}: '{target_sheet}'")

STATS_URL = os.getenv('STATS_URL', 'http://localhost:2145/')
FALLBACK_STATS_FILE = os.path.join(
    os.path.expanduser("~"),
    "Documents",
    "LethalCompanyStats",
    "stats.json"
)

START_ROW = 3

COLUMN_MAP = { #you can change the column where the data is put in (always starts in row 3 and it automatically goes onward)
    'MoonInfo_Name': 'F',
    'MoonInfo_Weather': 'G',
    'DungeonInfo_Interior': 'H',
    'DungeonInfo_ItemCount': 'I',
    'CollectedTotal': 'K',
    'BottomLine': 'L',
}


def parse_sse_payload(raw_text):
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    data_lines = []
    for line in lines:
        if line.startswith('data:'):
            data_lines.append(line[len('data:'):].strip())
    return '\n'.join(data_lines)


def get_stats_from_http():
    try:
        with urllib.request.urlopen(STATS_URL, timeout=5) as response:
            raw = response.read().decode('utf-8')

        raw = raw.lstrip('\ufeff')
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
        with open(FALLBACK_STATS_FILE, 'r', encoding='utf-8') as f:
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
    dungeon = stats.get('DungeonInfo') or {}
    moon = stats.get('MoonInfo') or {}
    return {
        'MoonInfo_Name': strip_apostrophe(moon.get('Name', '')),
        'MoonInfo_Weather': strip_apostrophe(moon.get('Weather', '')),
        'DungeonInfo_Interior': strip_apostrophe(dungeon.get('Interior', '')),
        'DungeonInfo_ItemCount': int(strip_apostrophe(dungeon.get('ItemCount', 0))),
        'CollectedTotal': int(strip_apostrophe(stats.get('CollectedTotal', 0))),
        'BottomLine': int(strip_apostrophe(stats.get('BottomLine', 0))),
    }

def get_next_empty_row():
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{target_sheet}!F{START_ROW}:F1000'
    ).execute()

    rows = result.get('values', [])
    next_row = START_ROW + len(rows)
    print(f"✓ Next empty row determined: {next_row}")
    return next_row


def write_to_cell(value, cell):
    try:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{target_sheet}!{cell}',
            valueInputOption='RAW',
            body={'values': [[value]]}
        ).execute()
        print(f"✓ Wrote '{value}' to {target_sheet}!{cell}")
    except Exception as e:
        print(f"✗ Error writing to {target_sheet}!{cell}: {e}")


def update_sheet_from_stats(stats):
    normalized = normalize_stats(stats)
    target_row = get_next_empty_row()

    for key, col in COLUMN_MAP.items():
        write_to_cell(normalized[key], f'{col}{target_row}')


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
                print('✓ No change in stats, skipping update')
        time.sleep(1)

if __name__ == '__main__':
    main()