import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager



#Convert date to time stamp
def timestmp(date_str):
    return int(datetime.datetime.strptime(date_str, "%Y-%m-%d").timestamp())


company = "NVDA"
start_time = timestmp("2026-06-24")
end_time = timestmp("2026-07-24")
url = f"https://finance.yahoo.com/quote/{company}/history/?period1={start_time}&period2={end_time}"

#Selenium setup&agent
options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    print("starting Yahoo Finance")
    driver.get(url)

    table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "table"))
    )

    # Get headers
    headers = [th.text.strip() for th in table.find_elements(By.XPATH, ".//thead//th") if th.text.strip()]

    # Get rows
    rows_data = []
    for row in table.find_elements(By.XPATH, ".//tbody//tr"):
        cells = [td.text.strip() for td in row.find_elements(By.TAG_NAME, "td")]
        if cells:
            rows_data.append(cells)

finally:
    driver.quit()


#Save to Excel & Auto-fit Columns
if rows_data:
    excel_file = f"{company}_history.xlsx"

    # Create DataFrame
    df = pd.DataFrame(rows_data, columns=headers if headers else None)

    # Save to Excel and adjust column widths cleanly
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]

        # Simple loop to set column widths based on max text length
        for col in ws.columns:
            col_letter = col[0].column_letter
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = max_len + 3

    print(f"Saved successfully {excel_file}")