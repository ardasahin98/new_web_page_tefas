import requests
import time
import os
import sys
import re
from lxml import html
import openpyxl

DEFAULT_TXT_FILE = "funds.txt"

if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

txt_file_path = os.path.join(base_dir, DEFAULT_TXT_FILE)

try:
    with open(txt_file_path, "r") as f:
        fonds = [
            line.strip().upper()
            for line in f
            if line.strip()
        ]
except Exception as e:
    print(f"Failed to read {txt_file_path}: {e}")
    sys.exit(1)

print(f"\nFunds to be processed: {fonds}\n")

file_path = os.path.join(base_dir, "tefas_funds.xlsx")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Funds"
ws.append(["Fund", "Price"])

session = requests.Session()

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

for fond_name in fonds:
    price = ""
    url = f"https://www.tefas.gov.tr/tr/fon-detayli-analiz/{fond_name}"

    try:
        response = session.get(url, headers=headers, timeout=20)
        print(f"{fond_name}: status code {response.status_code}")

        response.raise_for_status()

        tree = html.fromstring(response.content)

        elements = tree.xpath(
            "//p[contains(@class, 'font-bold') and contains(@class, 'text-primary-blue')]"
        )

        for el in elements:
            text = el.text_content().strip()
            if "," in text and any(char.isdigit() for char in text):
                price = text
                break

        if not price:
            page_text = response.text

            matches = re.findall(r"\d+,\d{4,}", page_text)
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

wb.save(file_path)

print(f"\nExcel file created:\n{file_path}")