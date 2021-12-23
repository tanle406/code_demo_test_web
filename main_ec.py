from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import win32com.client as win32
import pandas as pd
import datetime
import os
import sys
from time import sleep
import inspect
import logging
from shutil import copyfile


global LOG_INFO
CurDir = os.path.dirname(
    os.path.abspath(inspect.getfile(inspect.currentframe()))
)
logging.basicConfig(
    format="------------------------------- %(asctime)s >>>  %(message)s  <<<-------------------------------",
    datefmt="%d/%m/%Y %H:%M:%S",
)
logFormatter = logging.Formatter(
    "%(asctime)s ---------- %(message)s", datefmt="%d/%m/%Y %H:%M:%S"
)
LOG_INFO = logging.warning
FileHandler = logging.FileHandler("log.txt", "a+", "utf-8")
FileHandler.setFormatter(logFormatter)
logging.getLogger().addHandler(FileHandler)
LOG_INFO("-------------------------------START-------------------------------")

class Browser():
    def __init__(self, curdir, log_infor):
        self.curdir = curdir
        self.log_infor = log_infor
        self.list_result = []

    def open_browser_chrome(self):
        self.log_infor("Start open browser Chrome")
        chromeOptions = webdriver.ChromeOptions()
        driver = webdriver.Chrome(
            executable_path=os.path.abspath(self.curdir + "\\driver\\chromedriver.exe"),
            chrome_options=chromeOptions)
        self.log_infor("End open browser Chrome")
        return driver

    def wait_page_complete(self, driver, class_name):
        '''
        Check page complete and return the result
        class_name: class name need to check (string)
        '''
        delay = 30 # seconds
        try:
            myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, class_name)))
            result = True
            print("Page is ready!")
        except TimeoutException:
            result = False
            print("Loading took too much time!")
        return result

    def login_alert(self, user, password, driver):
        window_before = driver.window_handles[0]
        driver.switch_to_window(window_before)
        shell = win32.Dispatch("WScript.Shell")
        shell.Sendkeys(user)
        sleep(0.3)
        shell.Sendkeys('{TAB}')
        sleep(0.3)
        shell.Sendkeys(password)
        sleep(0.3)
        shell.Sendkeys('{TAB}')
        sleep(0.3)
        shell.Sendkeys('{ENTER}')

    def check_quick_order(self, driver):
        button_quick_order = driver.find_element(By.XPATH, '//*[@id="search"]/div/div[1]/span[2]/a/img')
        driver.execute_script("arguments[0].click();", button_quick_order)
        driver.find_element(By.ID, "0").send_keys("U110918")
        actions = ActionChains(driver)
        actions.send_keys(Keys.TAB * 1)
        actions.perform()
        self.wait_page_complete(driver, 'tdText0')
        str_prod = driver.find_element(By.CLASS_NAME, "tdText0").text
        if "い・ろ・は・す　ペットボトル　555ml　1ケース(※24本入)" in str_prod:
            return True
        else:
            return False

    def request_content(self, url):
        self.log_infor("----Start requests----")
        from bs4 import BeautifulSoup
        import requests
        content = requests.get(url, auth=("ec", "W7XxaowZ")).content
        soup = BeautifulSoup(content, 'html.parser')
        content_prod = soup.find("div", {"id": "product_wrapper"})
        self.log_infor("----End requests----")
        if "タイヤ" in content_prod.text:
            return True
        else:
            return False

    def search_tire(self, driver):
        arr_href_prod = []
        driver.find_element(By.CLASS_NAME, "search_input").send_keys("タイヤ")
        button_search = driver.find_element(By.ID, "btn_to_search")
        driver.execute_script("arguments[0].click();", button_search)
        sleep(2)
        list_prod = driver.find_elements(By.XPATH, '//*[@id="items_wrapper"]/div[*]/p/a')
        for item in list_prod:
            href_prod = item.get_attribute("href")
            arr_href_prod.append(href_prod)
        int_count = 0
        for href_prod in arr_href_prod:
            result = self.request_content(href_prod)
            if result:
                pass
            else:
                int_count += 1
        if int_count/len(arr_href_prod)*100 > 5:    # ratio percentage of product fail
            return False
        else:
            return True

    def manufacture_name(self, driver):
        driver.find_element(By.XPATH, '//*[@id="search"]/div/div[1]/span[3]/a/img').click()     # Click detail search
        driver.find_element(By.ID, "mk").send_keys("MICHELIN")
        driver.find_element(By.CLASS_NAME, "btn_red_search").click()    # Click search
        list_manufacture = driver.find_elements(By.XPATH, '//*[@id="items_wrapper"]/div[*]/div[1]/span')
        for manufacture_name in list_manufacture:
            if manufacture_name.text == "MICHELIN":
                pass
            else:
                return False
        return True

    def search_by_car(self, driver):
        sleep(2)
        btn_search_by_car = driver.find_element(By.XPATH, '//*[@id="more_posts"]/div[3]/div/ul/li[1]/div/a')
        driver.execute_script("arguments[0].click()", btn_search_by_car)    # Click button "Search By Car"

        btn_toyota = driver.find_element(By.LINK_TEXT, 'トヨタ')
        driver.execute_script("arguments[0].click()", btn_toyota)    # Click "TOYOTA"

        btn_a_line = driver.find_element(By.LINK_TEXT, 'アルファード')
        driver.execute_script("arguments[0].click()", btn_a_line)    # Click "A line"

        btn_model_car = driver.find_element(By.XPATH, '//*[@id="main"]/ul/li[3]/p[1]/a/img')
        driver.execute_script("arguments[0].click()", btn_model_car)    # Click "Model Car"

        btn_search_by_car = driver.find_element(By.XPATH, '//*[@id="main"]/table/tbody/tr[2]/td[4]/input')
        driver.execute_script("arguments[0].click()", btn_search_by_car)    # Click button "Search By Car"

        btn_ok = driver.find_element(By.XPATH, '//*[@id="default"]/div[21]/div/div[10]/button[1]')
        driver.execute_script("arguments[0].click()", btn_ok)    # Click button "Search By Car"
        self.wait_page_complete(driver, 'car_options')
        content = driver.find_element(By.CLASS_NAME, 'mb10').text
        if ("2009(平成21)年06月～2010(平成22)年04月" in content)\
            and "240S リミテッド 7人乗" in content:
            return True
        else:
            return False

    def search_by_part(self, driver):
        sleep(2)
        btn_search_by_part = driver.find_element(By.XPATH, '//*[@id="more_posts"]/div[3]/div/ul/li[2]/div/a')
        driver.execute_script("arguments[0].click()", btn_search_by_part)    # Click button "Search By Part"

        btn_wiper = driver.find_element(By.LINK_TEXT, 'ワイパー')
        driver.execute_script("arguments[0].click()", btn_wiper)    # Click button "Wiper"
        sleep(1)
        str_path = driver.find_element(By.CLASS_NAME, "nav_path").text
        str_title = driver.find_element(By.XPATH, '//*[@id="main"]/p').text
        if "ワイパー" in str_title and "ワイパー" in str_path:
            return True
        else:
            return False

    def main(self):
        url = "https://usamart-stg.a-it.jp/"
        driver = self.open_browser_chrome()
        driver.maximize_window()
        driver.get(url)
        sleep(3)
        self.login_alert("ec", "W7XxaowZ", driver)
        self.wait_page_complete(driver, "logo_div")
        # self.check_quick_order(driver)
        # result = self.search_tire(driver)
        # result = self.manufacture_name(driver)
        # result = self.search_by_car(driver)
        result = self.search_by_part(driver)
        print(result)

if __name__ == "__main__":
    curdir = os.path.dirname(
            os.path.abspath(inspect.getfile(inspect.currentframe())))
    driver = Browser(curdir, LOG_INFO)
    driver.main()
