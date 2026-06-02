import time
import os
import sys
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

DEFAULT_TXT_FILE = "funds.txt"

PRICE_XPATH = "/html/body/main/div[3]/div[2]/div[2]/div[1]/div[3]/div[2]/div[1]/div[2]/p"

if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

txt_file_path = os.path.join(base_dir, DEFAULT_TXT_FILE)
file_path = os.path.join(base_dir, "tefas_funds.xlsx")

try:
    with open(txt_file_path, "r", encoding="utf-8") as f:
        fonds = [line.strip().upper() for line in f if line.strip()]
except Exception as e:
    print(f"Failed to read {txt_file_path}: {e}")
    sys.exit(1)

print(f"\nFunds to be processed: {fonds}\n")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Funds"
ws.append(["Fund", "Price"])

options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1400,1000")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

for fond_name in fonds:
    price = ""

    url = f"https://www.tefas.gov.tr/tr/fon-detayli-analiz/{fond_name}"

    try:
        driver.get(url)

        element = wait.until(
            EC.presence_of_element_located((By.XPATH, PRICE_XPATH))
        )

        price = element.text.strip()

    except Exception as e:
        print(f"{fond_name}: price not found ({e})")

        debug_path = os.path.join(base_dir, f"debug_{fond_name}.html")
        with open(debug_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)

        print(f"{fond_name}: saved debug file: {debug_path}")

    ws.append([fond_name, price])
    print(f"{fond_name}: {price}")

    time.sleep(1)

driver.quit()
wb.save(file_path)

print(f"\nExcel file created:\n{file_path}")
