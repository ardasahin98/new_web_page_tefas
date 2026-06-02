import requests
import time
import os
import sys
from lxml import html
import openpyxl

DEFAULT_TXT_FILE = "funds.txt"

# ------------------ PATHS ------------------

if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

txt_file_path = os.path.join(base_dir, DEFAULT_TXT_FILE)

# ------------------ READ FUNDS ------------------

try:
    with open(txt_file_path, "r", encoding="utf-8") as f:
        fonds = [
            line.strip().upper()
            for line in f
            if line.strip()
        ]
except Exception as e:
    print(f"Failed to read {txt_file_path}: {e}")
    sys.exit(1)

print(f"\nFunds to be processed: {fonds}\n")

# ------------------ EXCEL FILE ------------------

file_path = os.path.join(base_dir, "tefas_funds.xlsx")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Funds"
ws.append(["Fund", "Price"])

# ------------------ HEADERS ------------------

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,tr;q=0.8",
    "Referer": "https://www.google.com/",
    "Connection": "keep-alive",
}

# ------------------ MAIN LOOP ------------------

for fond_name in fonds:
    price = ""

    url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={fond_name}"

    try:
        response = requests.get(url, headers=headers, timeout=10)

        print(f"{fond_name}: status code = {response.status_code}")
        print(f"{fond_name}: final url = {response.url}")

        response.raise_for_status()

        tree = html.fromstring(response.content)

        element = tree.xpath(
            "//*[@id='MainContent_PanelInfo']//ul/li[1]/span"
        )

        if element:
            price = element[0].text_content().strip()
        else:
            debug_path = os.path.join(base_dir, f"debug_{fond_name}.html")

            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(response.text)

            print(
                f"{fond_name}: price not found. "
                f"Saved debug file: {debug_path}"
            )

    except requests.exceptions.RequestException as e:
        print(f"{fond_name}: request failed ({e})")

    ws.append([fond_name, price])
    print(f"{fond_name}: {price}")

    time.sleep(1)

# ------------------ SAVE ------------------

wb.save(file_path)

print(f"\nExcel file created:\n{file_path}")
