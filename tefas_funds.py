import requests
import time
import os
import sys
import re
from lxml import html
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
        fonds = [
            line.strip().upper()
            for line in f
            if line.strip()
        ]
except Exception as e:
    print(f"Failed to read {txt_file_path}: {e}")
    sys.exit(1)

print(f"\nFunds to be processed: {fonds}\n")

# ------------------ EXCEL SETUP ------------------

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Funds"
ws.append(["Fund", "Price"])

# ------------------ HEADERS ------------------

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://www.tefas.gov.tr/",
    "Connection": "keep-alive",
}

session = requests.Session()

# ------------------ MAIN LOOP ------------------

for fond_name in fonds:
    price = ""

    url = f"https://www.tefas.gov.tr/tr/fon-detayli-analiz/{fond_name}"

    try:
        response = session.get(url, headers=headers, timeout=20)

        print(f"{fond_name}: status code = {response.status_code}")
        print(f"{fond_name}: final url = {response.url}")

        response.raise_for_status()

        tree = html.fromstring(response.content)

        # New TEFAS page XPath
        element = tree.xpath(
            "/html/body/main/div[3]/div[2]/div[2]/div[1]/div[3]/div[2]/div[1]/div[2]/p"
        )

        if element:
            price = element[0].text_content().strip()

        # Backup method
        if not price:
            elements = tree.xpath(
                "//p[contains(@class, 'font-bold') or contains(@class, 'text-primary-blue')]"
            )

            for el in elements:
                text = el.text_content().strip()
                if re.fullmatch(r"\d+,\d+", text):
                    price = text
                    break

        # Final backup regex
        if not price:
            matches = re.findall(r"\d+,\d{4,}", response.text)
            if matches:
                price = matches[0]

        if not price:
            debug_path = os.path.join(base_dir, f"debug_{fond_name}.html")
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(response.text)

            print(f"{fond_name}: price not found. Saved debug file: {debug_path}")

    except requests.exceptions.RequestException as e:
        print(f"{fond_name}: request failed ({e})")

    ws.append([fond_name, price])
    print(f"{fond_name}: {price}")

    time.sleep(1)

# ------------------ SAVE ------------------

wb.save(file_path)

print(f"\nExcel file created:\n{file_path}")
