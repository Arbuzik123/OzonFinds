import time
import re
import numpy as np
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from Searchengines.AddVal import add_value_to_next_empty_cell_in_row
from Searchengines.ConverExtract import convert_symbols_in_brackets

def SearchOzon(e, path, lock, X, Y, positions):
    options = Options()
    driver = uc.Chrome(options=options)
    driver.set_window_size(X, Y)
    driver.set_window_position(*positions, windowHandle='current')

    file_path = f'{path.split("_")[0]}_{e + 1}.xlsx'
    print(file_path)
    df = pd.read_excel(file_path)
    driver.get("https://www.ozon.ru/")
    time.sleep(3)
    lock.release()
    print("Старт функции")
    for index, row in df.iloc[:, 1].items():
        wordeee = row.split()
        url = "https://www.ozon.ru/search/?brand=87148725&from_global=true&text=" + "BRAIT " + wordeee[0] + " " + \
              df.iloc[index, 3] + "&from_global=true"
        driver.get(url)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        wait = WebDriverWait(driver, 3)
        try:
            print("Создаем цикл по элементам")
            elementozs = wait.until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@data-widget='searchResultsV2']/div/div")))
            elementozs = driver.find_elements(By.XPATH, "//div[@data-widget='searchResultsV2']/div/div")
            for element in elementozs:
                process_element(element, row, df, index, driver, path)
        except:
            try:
                driver.execute_script("window.localStorage.clear();")  # Очистить Local Storage
                driver.execute_script("window.sessionStorage.clear();")  # Очистить Session Storage
                driver.get(url)
                elementozs = wait.until(
                    EC.presence_of_all_elements_located((By.XPATH, "//div[@data-widget='searchResultsV2']/div/div")))
                elementozs = driver.find_elements(By.XPATH, "//div[@data-widget='searchResultsV2']/div/div")
                print("123")
                for element in elementozs:
                    process_element(element, row, df, index, driver, file_path)
            except:
                print("Нет элементов")
    df.to_excel(file_path, index=False)
    de = pd.read_excel(file_path)
    for index, row in de.iterrows():
        for column, value in row[4:].items():
            if pd.notna(value) and pd.notnull(value):
                print(value)
                driver.get(value)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                try:
                    new_link = driver.find_element(By.XPATH, "//div[@data-widget='webOutOfStock']//a").get_attribute("href")
                    driver.get(new_link)
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    print("Нет в наличии ссылка " + str(new_link))
                except:
                    print("В наличии")
                wait = WebDriverWait(driver, 5)
                try:
                    magazin = driver.find_element("xpath", '//div[@triggering-object-selector="#short-product-info-trigger-new"]/div/div[2]/div/div/div/a').get_attribute("title")
                except:
                    try:
                        magazin = driver.find_element("xpath", '//div[@triggering-object-selector="#short-product-info-trigger-new"]/div/div[2]/div[2]/div//span').get_attribute("title")
                    except:
                        magazin = driver.find_element("xpath",'//div[@data-widget="webCurrentSeller"]//div[2]/a').text
                try:
                    price = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-widget='webPrice']/div/div[2]/div/div/span"))).text
                except:
                    price = "0"

                try:
                    text = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-widget='webProductHeading']/h1"))).text
                except:
                    text = "unknown"
                try:
                    text1 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-widget='webDescription'][1]"))).text
                except:
                    text1 = "unknown"

                text = convert_symbols_in_brackets(text)
                text = re.sub(r'[^\w\s]', '', text).replace(" ", "").lower().replace("brait", "br")
                text1 = convert_symbols_in_brackets(text1)
                text1 = re.sub(r'[^\w\s]', '', text1).replace(" ", "").lower().replace("brait", "br")
                our_text = de.iloc[index, 3].replace(" ", "").lower().replace("brait", "")
                our_text = convert_symbols_in_brackets(our_text)
                our_text = re.sub(r'[^\w\s]', '', our_text)
                pattern = rf'{our_text}$|{our_text}(?![a-z])'
                similar_characters = {
                    'р': 'p',
                    'а': 'a',
                    'е': 'e',
                    'м': 'm',
                    'в': 'b',
                    'т': 't',
                    'с': 'c',
                    'н': 'h',
                    'о': 'o'
                }
                text = ''.join(similar_characters.get(char, char) for char in text)
                text1 = ''.join(similar_characters.get(char, char) for char in text1)
                matches = re.findall(pattern, text)
                matches2 = re.findall(pattern, text1)
                if matches or matches2:
                    print(f"Цена   {price}Магазин   {magazin} Валуе {value}")
                    print("Товар добавлен в таблицу")
                    new_data = {
                        'Наименование': de.iloc[index, 3],
                        'Store Name': magazin,
                        'Price': str(price).replace("₽", "").replace(" ", "").replace("c Ozon Картой", ""),
                        'Link': value,
                    }
                    row_index = de.index[de['Наименование'] == new_data['Наименование']]
                    if len(row_index) > 0:
                        store_col = f"{' '.join(new_data['Store Name'].split()).title()}"
                        if store_col in de.columns:
                            de.loc[row_index, store_col] = new_data['Price']
                            de.loc[row_index, f"{store_col} Link"] = new_data['Link']
                        else:
                            de[store_col] = np.nan
                            de.loc[row_index, store_col] = new_data['Price']
                            de.loc[row_index, f"{store_col} Link"] = new_data['Link']
                    else:
                        new_row = {
                            'Наименование': new_data['Наименование'],
                            'Store A': np.nan,
                            'Store B': np.nan,
                            'Store C': np.nan
                        }
                        store_col = f"{' '.join(new_data['Store Name'].split()).title()}"
                        new_row[store_col] = new_data['Price']
                        new_row[f"{store_col} Link"] = new_data['Link']
                        de = pd.concat([de, pd.DataFrame([new_row])], ignore_index=True)
                    de.to_excel(file_path, index=False)
                else:
                    print(f"Цена   {price}Магазин   {magazin} Валуе {value}")
                    print(our_text)
                    print(text)
                    print("Не подходит")
    print("Отправка файла")
    driver.close()
    driver.quit()

def process_element(element, row, df, index, driver, file_path):
    words = ['поршень', 'ремень', 'статор', 'шнек', 'трос', 'свеча', 'узел', 'регулятор', 'фильтр', 'реле',
             'стартер', 'шатун', 'катушка', 'карбюратор', 'подушки', 'подшипник', 'щетка', 'кран', 'мешок', 'мешки',
             'цепь', 'зарядное', 'шина', 'гайка', 'выключатель', 'рычаг', 'шаблон', 'мембрана', 'кожух', 'отвал', 'насадка', 'ротор', 'канат', 'головка', 'подушки']
    wait = WebDriverWait(driver, 2)
    wait.until(EC.presence_of_element_located((By.XPATH, ".//div/div/a")))
    link = element.find_element(By.XPATH, ".//div/div/a")
    text = link.text.lower()
    our_text = df.iloc[index, 1].split()[0].lower()
    if our_text in text:
        if not any(word in text for word in words):
            print(our_text + " " + text)
            print("Слово подходит")
            add_value_to_next_empty_cell_in_row(df, index, str(link.get_attribute("href")))
            df.to_excel(file_path, index=False)
            print("Успешно добавлено")

def remove_after_lowercase(input_string):
    for i, char in enumerate(input_string):
        if char.islower():
            return input_string[:i]
    return input_string
