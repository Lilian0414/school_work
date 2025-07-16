import pandas as pd
import re
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 讀取 Excel
df = pd.read_excel("excessive_thesis.xlsx")
df_to_check = df[df["備註"].isna()].copy()

# 設定 Selenium
options = webdriver.ChromeOptions()
# options.add_argument("--headless")  # 若想看畫面可註解掉
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 15)

for idx, row in df_to_check.iterrows():
    name = str(row["name"]).strip()
    result = "查無"
    print(f"\n🔍 正在查詢：{name}")

    try:
        # 進入首頁
        driver.get("https://ndltd.ncl.edu.tw/")

        # 打勾指導教授（注意 name="dcf" value="ad"）
        checkbox = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@type='checkbox' and @name='dcf' and @value='ad']")))
        if not checkbox.is_selected():
            checkbox.click()

        # 輸入教授名（注意 name="qs0"）
        search_box = wait.until(EC.visibility_of_element_located((By.NAME, "qs0")))
        search_box.clear()
        search_box.send_keys(name)
        time.sleep(0.3)

        # 點查詢按鈕（圖片按鈕 name="gs32search"）
        search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "gs32search")))
        search_btn.click()

        # 等載入結果頁面
        time.sleep(2)

        # 抓 class="etd_e" 的所有 span，找出純數字（即論文數量）
        spans = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "etd_e")))
        numbers = [s.text.strip() for s in spans if s.text.strip().isdigit()]
        if numbers:
            result = numbers[0]
        else:
            print("⚠️ 找不到筆數")

    except Exception as e:
        print(f"❌【{name}】查詢失敗：{e}")
        print("🕵️ HTML 頭 500 字：", driver.page_source[:500])

    print(f"✅ {name} → {result}")
    df.at[idx, "備註"] = result
    time.sleep(random.uniform(1.5, 3.5))  # 隨機等待

driver.quit()
df.to_excel("excessive_thesis_checked.xlsx", index=False)
print("\n✅ 所有查詢完成！結果儲存為 excessive_thesis_checked111.xlsx")
