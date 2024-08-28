import time
import re
import pandas as pd
import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
def updateOzon(e, path, lock, X, Y, positions):
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=C:/Users/User/AppData/Local/Google/Chrome/User Data")
    options.add_argument('--profile-directory=Profile 1')
    driver = uc.Chrome(options=options)
    driver.set_window_size(X, Y)
    driver.set_window_position(*positions, windowHandle='current')
    file_path = f'{path.split("_")[0]}_{e + 1}.xlsx'
    df = pd.read_excel(file_path)
    driver.get("https://seller.ozon.ru/app/brand-products/all")
    time.sleep(15)
    wait = WebDriverWait(driver, 5)
    for col_name in df.columns[4:]:
        for index, value in df[col_name].items():
            if pd.notna(value) and not pd.api.types.is_numeric_dtype(value) and str(value).startswith('http'):
                if str(value).startswith('https://www.ozon.ru/'):
                    try:
                        match = re.search(r'-(\d+)(?=/\?asb=)|-(\d+)(?=/\?advert=)', value)
                        art = match.group(1) if match else re.search(r'[^-]+$', value).group(0)
                        print(art)

                        element = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']")))
                        driver.find_element(By.XPATH, "//input[@type='text']").clear()
                        driver.find_element(By.XPATH, "//input[@type='text']").send_keys(art)
                        time.sleep(1)
                    except:
                        print("hz che")
                    try:
                        driver.find_element(By.XPATH, "//td[.=' Нет записей ']")
                        price = "НЕ НАШ БРЕНД"
                    except:
                        time.sleep(1)
                        try:

                            element = wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr/td[7]")))
                            price = driver.find_element(By.XPATH, "//tbody/tr/td[7]").text
                            prev_col_index = df.columns.get_loc(col_name) - 1
                            prev_col_name = df.columns[prev_col_index]
                            df.at[index, prev_col_name] = price
                        except:
                            price = "НЕ НАШ БРЕНД"
                            prev_col_index = df.columns.get_loc(col_name) - 1
                            prev_col_name = df.columns[prev_col_index]
                            df.at[index, prev_col_name] = price

                    print(price)
    # third_col_name = df.columns[2]
    # # Итерация по строкам третьего столбца
    # for index, value in df[third_col_name].items():
    #     try:
    #         # match = re.search(r'-(\d+)(?=/\?asb=)|-(\d+)(?=/\?advert=)', value)
    #         # art = match.group(1) if match else re.search(r'[^-]+$', value).group(0)
    #         # print(art)
    #
    #         element = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']")))
    #         driver.find_element(By.XPATH, "//input[@type='text']").clear()
    #         driver.find_element(By.XPATH, "//input[@type='text']").send_keys(value)
    #         time.sleep(1)
    #     except:
    #         print("hz che")
    #     try:
    #         driver.find_element(By.XPATH, "//td[.=' Нет записей ']")
    #         price = "НЕ НАШ БРЕНД"
    #     except:
    #         time.sleep(1)
    #         try:
    #             element = wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr/td[7]")))
    #             price = driver.find_element(By.XPATH, "//tbody/tr/td[7]").text
    #             prev_col_index = df.columns.get_loc(third_col_name) + 3
    #             prev_col_name = df.columns[prev_col_index]
    #             link = driver.find_element(By.XPATH,"//tbody/tr/td[2]//a").get_attribute("href")
    #             df.at[index, prev_col_name] = price
    #             prev_col_index1 = df.columns.get_loc(third_col_name) + 4
    #             prev_col_name1 = prev_col_index1 + 1
    #             df.at[index, prev_col_name1] = link
    #         except:
    #             price = "Нет товара"
    #             prev_col_index = df.columns.get_loc(third_col_name) + 3
    #             prev_col_name = df.columns[prev_col_index]
    #             df.at[index, prev_col_name] = price

        # print(price)
        df.to_excel("itog.xlsx",index=False)
        df.to_excel(file_path, index=False)
    driver.quit()