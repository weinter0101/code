# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 13:52:09 2024
@author: USER
"""
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from selenium.webdriver.chrome.options import Options

def safe_click(driver, selector, by=By.CSS_SELECTOR, max_attempts=3):
    for attempt in range(max_attempts):
        try:
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((by, selector))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            element.click()
            return True
        except (ElementClickInterceptedException, TimeoutException) as e:
            print(f"嘗試 {attempt + 1}: 點擊失敗. 錯誤: {str(e)}")
            try:
                driver.execute_script("arguments[0].click();", element)
                return True
            except Exception:
                print("JavaScript 點擊也失敗")
            time.sleep(2)
    print(f"無法點擊元素: {selector}")
    return False

chrome_options = Options()
chrome_options.add_argument("--disable-notifications")
driver = webdriver.Chrome(options=chrome_options)

driver.get('https://www.sinyi.com.tw/tradeinfo/list')
time.sleep(2)

# city
safe_click(driver, '//*[@id="citySingleAreaSelector CitySingleAreaWeb"]/div[1]/div', By.XPATH)
cityXPATH = '//*[@id="citySingleAreaSelector CitySingleAreaWeb"]/div[3]/div/div[2]/div[1]/label/span'
safe_click(driver, cityXPATH, By.XPATH)

# time
safe_click(driver, '//*[@id="area-and-community-menu-bottom"]/div[4]/div[1]/div', By.XPATH)
timeXPATH = '//*[@id="_opts"]/div[6]'
safe_click(driver, timeXPATH, By.XPATH)

# type
safe_click(driver, '//*[@id="area-and-community-menu-bottom"]/div[5]/div/div[1]', By.XPATH)
safe_click(driver, '//*[@id="area-and-community-menu-bottom"]/div[5]/div[2]/div[2]/div/div[1]/div/label/div', By.XPATH)
safe_click(driver, '//*[@id="area-and-community-menu-bottom"]/div[5]/div[2]/div[2]/div/div[2]/div/label/div', By.XPATH)
safe_click(driver, '//*[@id="area-and-community-menu-bottom"]/div[5]/div[2]/div[2]/div/div[3]/div/label/div', By.XPATH)

safe_click(driver, '//*[@id="__next"]/div', By.XPATH)
safe_click(driver, '//*[@id="area-and-community-menu-bottom"]/div[7]/div', By.XPATH)
time.sleep(1)

# 拿資料
js_script = """
return Array.from(document.querySelectorAll('.trade-obj-card-web')).map(el => {
    const year = el.querySelector('span:not([style])').textContent;
    const address = el.querySelector('div[style="width: calc(17% + 1px); overflow-y: auto;"] span').textContent;
    return {year, addreˇss};
});
"""

all_dataframes = []
results = driver.execute_script(js_script)
all_dataframes.append(pd.DataFrame(results))

# 每一頁
for page in range(2, 1718):
    css_selector = f'.pageClassName a[aria-label^="Page {page}"]'
    if safe_click(driver, css_selector):
        print(f"成功點擊頁面 {page}")
        results = driver.execute_script(js_script)
        all_dataframes.append(pd.DataFrame(results))
        time.sleep(2)
    else:
        print(f"無法處理第 {page} 頁")

# 合併
if all_dataframes:
    data = pd.concat(all_dataframes, ignore_index=True)
    # 儲存資料
    data.to_csv('房地產資料.csv', index=False, encoding='utf-8-sig')
    print(f"共收集到 {len(data)} 筆資料")
else:
    print("沒有成功獲取任何資料")

# 關閉瀏覽器
driver.quit()

data.to_excel('C:/Users/USER/Downloads/outputt.xlsx')