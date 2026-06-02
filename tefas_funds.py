import requests
import time
import os
import sys
import openpyxl

DEFAULT_TXT_FILE = "funds.txt"

# ------------------ PATHS ------------------

if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

txt_file_path = os.path.join(base_dir, DEFAULT_TXT_FILE)
file_path = os.path.join(base_dir, "tefas_funds.xlsx")

# ------------------ READ FUNDS ------------------

try:
    with open(txt_file_path, "r", encoding="utf-8") as f:
        fonds = [line.strip().upper() for line in f if line.strip()]
except Exception as e:
    print(f"Failed to read {txt_file_path}: {e}")
    sys.exit(1)

print(f"\nFunds to be processed: {fonds}\n")

# ------------------ EXCEL SETUP ------------------

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Funds"
ws.append(["Fund", "Price"])

# ------------------ API SETUP ------------------

url = "https://www.tefas.gov.tr/api/funds/fonFiyatBilgiGetir"

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
}

session = requests.Session()
session.headers.update(headers)

# ------------------ MAIN LOOP ------------------

for fond_name in fonds:
    price = ""

    payload = {
        "fonKodu": fond_name,
        "dil": "TR",
        "periyod": 1
    }

    try:
        response = session.post(url, json=payload, timeout=30)
        print(f"{fond_name}: status code = {response.status_code}")

        response.raise_for_status()

        data = response.json()
        rows = data.get("resultList", [])

        if rows:
            latest_row = rows[-1]
            price = latest_row.get("fiyat", "")

            if price == "":
                price = latest_row.get("price", "")

        if price == "":
            print(f"{fond_name}: price not found")
            print(data)

    except Exception as e:
        print(f"{fond_name}: request failed ({e})")

    ws.append([fond_name, price])
    print(f"{fond_name}: {price}")

    time.sleep(1)

# ------------------ SAVE ------------------

wb.save(file_path)

print(f"\nExcel file created:\n{file_path}")
