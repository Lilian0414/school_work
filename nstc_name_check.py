import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 載入名單
df = pd.read_excel("excessive_thesis.xlsx")
df_to_check = df[df["備註"].isna()].copy()

# 啟動瀏覽器
options = webdriver.ChromeOptions()
# options.add_argument("--headless")  # 可選：關掉這行就會跳出畫面
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 15)

for idx, row in df_to_check.iterrows():
    name = str(row["name"]).strip()
    is_duplicate = "查詢失敗"

    try:
        # 開啟查詢頁面
        driver.get("https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initSearchList&LANG=chi")

        # 找到姓名輸入欄位，填入姓名
        name_input = wait.until(EC.visibility_of_element_located((By.ID, "nameChi")))
        name_input.clear()
        name_input.send_keys(name)

        # 點擊查詢按鈕
        search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='查詢']")))
        search_btn.click()

        # 等待結果表格顯示，統計筆數（每位教授一列 <tr>）
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "page")))
        time.sleep(2)  # 有時需要等 AJAX

        rows = driver.find_elements(By.XPATH, "//table[contains(@class, 'C30Tblist2')]/tbody/tr")
        visible_rows = [r for r in rows if r.text.strip() != ""]

        if len(visible_rows) > 1:
            is_duplicate = "是"
        elif len(visible_rows) == 1:
            is_duplicate = "否"
        else:
            is_duplicate = "查無"

    except Exception as e:
        print(f"❌ 查詢錯誤：{name} → {e}")
        is_duplicate = "查詢錯誤"

    df.at[idx, "是否重複"] = is_duplicate
    print(f"✅ {name} → {is_duplicate}")
    time.sleep(2)

driver.quit()
df.to_excel("excessive_thesis_checked_duplicates.xlsx", index=False)
print("✅ 重複名冊查詢完成，已儲存為 excessive_thesis_checked_duplicates.xlsx")

