from appium import webdriver
from appium.webdriver.common.appiumby import AppiumBy
from appium.webdriver.common.touch_action import TouchAction
import win32com.client as win32
import pandas as pd
import datetime
import os
import sys
import base64
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


class Android():
    def __init__(self, curdir, log_infor):
        self.curdir = curdir
        self.log_infor = log_infor
        self.list_result = []

    def connect_android(self):
        desired_cap = {
            "deviceName": "Pixel 5 API 27",
            "platformName": "Android",
        }

        driver = webdriver.Remote('http://localhost:4723/wd/hub', desired_cap)
        driver.implicitly_wait(30)
        return driver

    def open_chrome_and_go_to_url(self, url, driver):
        driver.find_element(AppiumBy.ACCESSIBILITY_ID ,"Chrome").click()
        sleep(3)
        element_type = driver.find_element(AppiumBy.ID ,"com.android.chrome:id/search_box_text")
        element_type.send_keys(url)
        driver.press_keycode(66) # Press ENTER
        sleep(2)

    def create_folder(self, path_folder):
        check_folder = os.path.exists(path_folder)
        if check_folder:
            pass
        else:
            os.makedirs(path_folder)

    def move_file(self, old_path, new_location):
        time_now = datetime.datetime.now()
        name_file = old_path.split("\\")[-1]
        new_path = new_location + "\\" + name_file
        if os.path.exists(old_path):
            if not os.path.exists(new_path):
                os.rename(old_path, new_path)
            elif os.path.exists(new_path):
                new_path = new_location + "\\" + \
                    time_now.strftime("%Y%m%d%f") + "_" + name_file
                os.rename(old_path, new_path)

    def scroll_down(self, driver, n):
        for i in range(n):
            touch = TouchAction(driver)
            touch.long_press(x=567, y=1726).move_to(x=567, y=488).release().perform()   # Scroll down
            sleep(2)

    def scroll_up(self, driver, n):
        for i in range(n):
            touch = TouchAction(driver)
            touch.long_press(x=567, y=488).move_to(x=567, y=1726).release().perform()   # Scroll up
            sleep(2)

    def check_button_page_top(self, driver):
        '''
        Check button "Page Top" and return the result
        '''
        self.log_infor("Start check button Page Top")
        self.scroll_down(driver, 2)
        TouchAction(driver).tap(x=993, y=1703).perform()
        try:
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="アヲハタ フルーツには続きがある。"]')
            return True
        except:
            return False

    def get_url(self, driver):
        return driver.find_element(AppiumBy.ID, 'com.android.chrome:id/url_bar').text

    def check_logo_home_page_above(self, driver):
        '''
        Check home page and return the result
        '''
        self.log_infor("Check button Logo Home Page Above")
        driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="アヲハタ フルーツには続きがある。"]').click()
        sleep(3)
        str_url = self.get_url(driver)
        self.scroll_down(driver, 1)
        result = True
        try:
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="商品情報 "]')   # check element home page
        except:
            result = False
        self.scroll_up(driver, 1)
        if str_url == "https://www.aohata.co.jp" and result:
            return True
        else:
            return False

    def check_page_recommend_products(self, driver):
        '''
        Check button Recommend Products and return the result
        '''
        self.log_infor("Start check page Recommend Products")
        try:
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="#"]').click()
            sleep(1)
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="商品情報"]').click()
            sleep(3)
            str_url = self.get_url(driver)
            self.scroll_down(driver, 1)
            try:
                driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="シリーズラインアップ "]')
                result = True
            except:
                result = False
            self.scroll_up(driver, 1)
            if "https://www.aohata.co.jp/products/" in str_url and result:
                return True
            else:
                return False
        except Exception as error:
            print(error)
            return False

    def check_page_recommend_recipe(self, driver):
        '''
        Check button Recommend Recipe and return the result
        '''
        self.log_infor("Start check page Recommend Recipe")
        try:
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="#"]').click()
            sleep(1)
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="おすすめレシピ"]').click()
            sleep(3)
            str_url = self.get_url(driver)
            self.scroll_down(driver, 1)
            try:
                driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="レシピを絞り込む"]')
                result = True
            except:
                result = False
            self.scroll_up(driver, 1)
            if "https://www.aohata.co.jp/recipes/" in str_url and result:
                return True
            else:
                return False
        except Exception as error:
            print(error)
            return False

    def check_page_experience(self, driver):
        '''
        Check button Experience and return the result
        '''
        self.log_infor("Start check page Experience")
        try:
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="#"]').click()
            sleep(1)
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="知る・見る・体験する"]').click()
            sleep(3)
            str_url = self.get_url(driver)
            self.scroll_down(driver, 1)
            try:
                driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="詳しくはこちら"]')
                result = True
            except:
                result = False
            self.scroll_up(driver, 1)
            if "https://www.aohata.co.jp/experience/" in str_url and result:
                return True
            else:
                return False
        except Exception as error:
            print(error)
            return False

    def check_page_company(self, driver):
        '''
        Check button Company and return the result
        '''
        self.log_infor("Start check page Company")
        return False
        # try:
        #     driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="#"]').click()
        #     sleep(1)
        #     driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="企業情報"]').click()
        #     sleep(3)
        #     str_url = self.get_url()
        #     self.scroll_down(driver, 1)
        #     try:
        #         driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="詳しくはこちら"]')
        #         result = True
        #     except:
        #         result = False
        #     if "https://www.aohata.co.jp/company/" in str_url and result:
        #         return True
        #     else:
        #         return False
        # except:
        #     return False

    def check_page_contact(self, driver):
        '''
        Check button Contact and return the result
        '''
        self.log_infor("Start check page Contact")
        try:
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="#"]').click()
            sleep(1)
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="お問い合わせ・FAQ"]').click()
            sleep(3)
            str_url = self.get_url(driver)
            try:
                driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="よくあるご質問 "]')
                result = True
            except:
                result = False
            if "https://www.aohata.co.jp/inquiry/" in str_url and result:
                return True
            else:
                return False
        except Exception as error:
            print(error)
            return False

    def check_page_recruitment(self, driver):
        '''
        Check button Recruitment and return the result
        '''
        self.log_infor("Start check page Recruitment")
        try:
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="#"]').click()
            sleep(1)
            driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="採用情報"]').click()
            sleep(3)
            str_url = self.get_url(driver)
            self.scroll_down(driver, 1)
            try:
                driver.find_element(AppiumBy.XPATH, '//android.view.View[@text="入社後の待遇"]')
                result = True
            except:
                result = False
            self.scroll_up(driver, 1)
            if "https://www.aohata.co.jp/recruit.html" in str_url and result:
                return True
            else:
                return False
        except Exception as error:
            print(error)
            return False

    def check_button_english(self, driver):
        '''
        Check button English and return the result
        '''
        self.log_infor("Start check button English")

    def check_button_chinese(self, driver):
        '''
        Check button Chinese and return the result
        '''
        self.log_infor("Start check button Chinese")

    def check_button_next_slick_arrow(self, driver):
        '''
        Check button next slick arrow and return the result
        '''
        self.log_infor("Start check button next slick arrow")
        sleep(2)
        text = driver.find_element(AppiumBy.XPATH, '//android.view.View[@bounds="[0,363][1080,1402]"]').get_attribute('text')
        print(text)
        TouchAction(driver).tap(x=1012, y=1328).perform()
        sleep(1)
        text = driver.find_element(AppiumBy.XPATH, '//android.view.View[@bounds="[0,363][1080,1402]"]').get_attribute('text')


    def check_button_prev_slick_arrow(self, driver):
        '''
        Check button next slick arrow and return the result
        '''
        self.log_infor("Start check button prev slick arrow")
        sleep(2)
        text = driver.find_element(AppiumBy.XPATH, '//android.view.View[@bounds="[0,351][1080,1393]"]').get_attribute('text')
        TouchAction(driver).tap(x=70, y=1328).perform()
        sleep(1)
        text = driver.find_element(AppiumBy.XPATH, '//android.view.View[@bounds="[0,351][1080,1393]"]').get_attribute('text')


    def write_to_excel(self, output_file, col_report, col_date):
        '''
        write report to excel file with column report and column date
        col_report, col_date: Eg: I, K, ...
        '''
        from openpyxl import load_workbook
        time_now = datetime.datetime.now().strftime("%m/%d")
        work_book = load_workbook(output_file)
        work_sheet = work_book.active
        int_row = 10
        for result in self.list_result:
            if result:
                work_sheet[col_report + str(int_row)].value = "OK"
            else:
                work_sheet[col_report + str(int_row)].value = "NG"
            work_sheet[col_date + str(int_row)].value = time_now
            int_row += 1
        work_book.save(output_file)


    def backup_and_create_output_file(self):
        '''
        move the output file into backup folder and create a new output file into the output folder
        '''
        self.log_infor("Start backup and create a new output file")
        output_file = self.curdir + "\\output\\output_android.xlsx"
        backup_directory = self.curdir + "\\bak\\output_android_" + datetime.datetime.now().strftime("%Y%m%d")
        self.create_folder(backup_directory)    # create a backup folder by day
        self.move_file(output_file, backup_directory)   # move the output file into the backup folder

        copyfile("D:\\TanLV\\SS1\\tmpl\\tmpl_output_android.xlsx", output_file)   # create a new output file by tmpl file
        self.log_infor("End backup and create a new output file")
        return output_file

    def main_android(self, output_file):
        driver = self.connect_android()
        self.open_chrome_and_go_to_url("https://www.aohata.co.jp", driver)
        driver.start_recording_screen()     # Start record video
        result = self.check_logo_home_page_above(driver)
        self.list_result.append(result)
        result = self.check_page_recommend_products(driver)
        self.list_result.append(result)

        result = self.check_page_recommend_recipe(driver)
        self.list_result.append(result)

        result = self.check_page_experience(driver)
        self.list_result.append(result)

        self.list_result.append(False)  # Company

        result = self.check_page_contact(driver)
        self.list_result.append(result)

        result = self.check_page_recruitment(driver)
        self.list_result.append(result)

        self.list_result.append(False)  # English
        self.list_result.append(False)  # Chinese

        self.check_logo_home_page_above(driver)
        self.list_result.append(False)  # Next slide
        self.list_result.append(False)  # Previous slide
        result = self.check_button_page_top(driver)
        self.list_result.append(result)
        self.write_to_excel(output_file, "I", "K")
        video_craw = driver.stop_recording_screen()     # Stop record video
        file_path_video = "D:\\TanLV\\SS1\\video\\video_record_android.mp4"
        with open(file_path_video, "wb") as file:
            file.write(base64.b64decode(video_craw))
        driver.quit()


if __name__ == "__main__":
    curdir = os.path.dirname(
            os.path.abspath(inspect.getfile(inspect.currentframe())))
    driver = Android(curdir, LOG_INFO)
    output_file = driver.backup_and_create_output_file()
    driver.main_android(output_file)