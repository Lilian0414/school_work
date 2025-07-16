from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

prof_name = ""

options = webdriver.ChromeOptions()
# 不加 headless，這樣你能看到畫面
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 15)

driver.get("https://ndltd.ncl.edu.tw/")

# ✅ 打勾「指導教授」
checkbox = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//input[@type='checkbox' and @name='dcf' and @value='ad']")))
if not checkbox.is_selected():
    checkbox.click()
print("✅ 已勾選『指導教授』")

# ✅ 輸入教授姓名
search_box = wait.until(EC.visibility_of_element_located((By.NAME, "qs0")))
search_box.clear()
search_box.send_keys(prof_name)
time.sleep(0.3)

# ✅ 點擊查詢按鈕（圖片按鈕）
search_button = wait.until(EC.element_to_be_clickable((By.NAME, "gs32search")))
search_button.click()

# ✅ 等結果顯示
time.sleep(2)
brief = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "brief")))
print("📄 查詢結果文字：", brief.text)

# ✅ 給你看畫面
time.sleep(10)
driver.quit()
