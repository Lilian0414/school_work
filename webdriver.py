import pandas as pd
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 讀取 Excel
df = pd.read_excel("excessive_thesis.xlsx")
df_to_check = df[df["備註"].isna()].copy()

# Selenium 設定
options = webdriver.ChromeOptions()
#options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 10)

for idx, row in df_to_check.iterrows():
    name = str(row["name"]).strip()
    result = "查無"

    try:
        driver.get("https://ndltd.ncl.edu.tw/")

        # 打勾「指導教授」checkbox
        advisor_checkbox = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='f1' and @value='指導教授']")))
        if not advisor_checkbox.is_selected():
            advisor_checkbox.click()

        # 輸入姓名
        search_box = driver.find_element(By.NAME, "q")
        search_box.clear()
        search_box.send_keys(name)

        # 點查詢按鈕
        driver.find_element(By.XPATH, "//input[@value='查詢']").click()

        # 等待結果
        time.sleep(2)
        result_text = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "brief"))).text
        print(f"📄 原始搜尋結果文字：{result_text}")


        match = re.search(r"共\s*(\d+)\s*筆", result_text)
        if match:
            result = match.group(1)

    except Exception as e:
        print(f"❌ {name} 查詢錯誤：{e}")

    print(f"✅ {name} → {result}")
    df.at[idx, "備註"] = result

driver.quit()
df.to_excel("excessive_thesis_checked.xlsx", index=False)


