from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import pandas as pd
import numpy as np
import until
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
        self.results = []
        self.values = []
        self.images_name = []
        self.int_count = 1

    def download_image(self, image_url, dir_folder):
        import requests
        import shutil
        filename = os.path.basename(image_url)
        path_image = f"{dir_folder}\\{filename}"
        try:
            r = requests.get(image_url, stream=True)
            if r.status_code == 200:
                r.raw.decode_content = True
                with open(path_image, 'wb') as f:
                    shutil.copyfileobj(r.raw, f)
                self.log_infor('Download successfull: ' + str(image_url))
            else:
                self.log_infor('Download fail: ' + str(image_url))
        except Exception as error:
            self.log_infor(
                "Error on line " + str(sys.exc_info()[-1].tb_lineno) + " " + type(error).__name__ + " " + str(error)
            )
        return path_image

    def compare_images(self, path_image_original, path_image_banner):
        import cv2
        import numpy as np
        try:
            image1 = cv2.imread(path_image_original)
            image2 = cv2.imread(path_image_banner)
            diff = cv2.subtract(image1, image2)
            if not np.any(diff):
                return True
            else:
                return False
        except:
            return False

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

    def DisplayMessageBox(self, body, title="Message", type="info"):
        '''
        Shows a pop-up message with title and body. Three possible types, info, error and warning with the default value info.
        '''
        import tkinter
        from tkinter import messagebox

        # hide main window
        root = tkinter.Tk()
        root.withdraw()
        if not body:
            messagebox.showwarning("Warning", "No input for message box")

        if type == "error":
            messagebox.showwarning(title, body)
        if type == "warning":
            messagebox.showwarning(title, body)
        if type == "info":
            messagebox.showinfo(title, body)
        return

    def open_browser_chrome(self):
        self.log_infor("Start open browser Chrome")
        chromeOptions = webdriver.ChromeOptions()
        driver = webdriver.Chrome(
            executable_path=os.path.abspath(f"{self.curdir}\\driver\\chromedriver.exe"),
            chrome_options=chromeOptions)
        self.log_infor("End open browser Chrome")
        return driver

    def go_to_url(self, url, driver):
        '''
        go to website
        '''
        self.log_infor("Start go to website")
        driver.get(url)
        self.log_infor("End go to website")

    def wait_page_complete(self, driver, class_name):
        '''
        Check page complete and return the result
        class_name: class name need to check (string)
        '''
        delay = 10 # seconds
        try:
            WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, class_name)))
            result = True
            print("Page is ready!")
        except TimeoutException:
            result = False
            print("Loading took too much time!")
        return result

    def check_slider_list_item(self, driver):
        '''
        check slider list item and return the result
        '''
        self.log_infor("Start check slider list item")
        try:
            driver.find_element(By.CLASS_NAME, 'c-sliderlist__item')
            result = True
        except Exception as error:
            result = False
        self.log_infor("Result check slider list item: {}".format(result))
        self.log_infor("End check slider list item")
        return result

    def screen_shot(self, driver, image_name, result):
        image_name = self.curdir + "\\image\\" + str(self.int_count) + "-" + image_name + "-" + str(result) + ".png"
        image_name = image_name.replace("True", "OK").replace("False", "NG")
        sleep(1)
        until.fullpage_screenshot(driver, image_name)
        self.int_count += 1
        return image_name

    def check_logo_home_page_above(self, driver, df):
        '''
        Check home page and return the result
        '''
        self.log_infor("Start check button Logo Home Page Above")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "???????????????????????????????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, 'c-sliderlist__item')
            result_check_slider_item = self.check_slider_list_item(driver)
            title = driver.title
            url = driver.current_url
            if ('??????????????????????????????' in str(title)) and result_check_slider_item:
                result = True
            else:
                result = False
        except Exception as error:
            self.log_infor(
                "Error on line " + str(sys.exc_info()[-1].tb_lineno) + " " + type(error).__name__ + " " + str(error)
            )
            result = False
        image = self.screen_shot(driver, "???????????????????????????????????????????????????", result)
        self.log_infor("Result check button Logo Home Page Above: {}".format(result))
        self.log_infor("End check button Logo Home Page Above")
        return result, url, image

    def check_page_recommend_products(self, driver, df):
        '''
        Check button Recommend Products and return the result
        '''
        self.log_infor("Start check page Recommend Products")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "???????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, 'c-hdg--lv1')
            title = driver.title
            url = driver.current_url
            if "????????????" in title:
                result = True
            else:
                result = False
        except Exception as error:
            self.log_infor(
                "Error on line " + str(sys.exc_info()[-1].tb_lineno) + " " + type(error).__name__ + " " + str(error)
            )
            result = False
        image = self.screen_shot(driver, "???????????????????????????", result)
        print(image)
        self.log_infor("Result check page Recommend Products: {}".format(result))
        self.log_infor("End check page Recommend Products")
        return result, url, image

    def check_page_recommend_recipe(self, driver, df):
        '''
        Check button Recommend Recipe and return the result
        '''
        self.log_infor("Start check page Recommend Recipe")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "????????????????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, 'p-rcp__search')
            title = driver.title
            url = driver.current_url
            if "?????????????????????" in title:
                result = True
            else:
                result = False
        except Exception as error:
            result = False
        image = self.screen_shot(driver, "????????????????????????????????????", result)
        self.log_infor("Result check page Recommend Recipe: {}".format(result))
        self.log_infor("End check page Recommend Recipe")
        return result, url, image

    def check_page_experience(self, driver, df):
        '''
        Check button Experience and return the result
        '''
        self.log_infor("Start check page Experience")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "?????????????????????????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, 'p__experience__news__block')
            title = driver.title
            url = driver.current_url
            if "??????????????????????????????" in title:
                result = True
            else:
                result = False
        except Exception as error:
            result = False
        image = self.screen_shot(driver, "?????????????????????????????????????????????", result)
        self.log_infor("Result check page Experience: {}".format(result))
        self.log_infor("End check page Experience")
        return result, url, image

    def check_page_company(self, driver, df):
        '''
        Check button Company and return the result
        '''
        self.log_infor("Start check page Company")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "???????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, 'l-info-header__inr')
            title = driver.title
            url = driver.current_url
            if "????????????" in title:
                result = True
            else:
                result = False
        except Exception as error:
            result = False
        image = self.screen_shot(driver, "???????????????????????????", result)
        self.log_infor("Result check page Company: {}".format(result))
        self.log_infor("End check page Company")
        return result, url, image

    def check_page_contact(self, driver, df):
        '''
        Check button Contact and return the result
        '''
        self.log_infor("Start check page Contact")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "?????????????????????FAQ???????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, 'l-inquiry')
            title = driver.title
            url = driver.current_url
            if "?????????????????????FAQ" in title:
                result = True
            else:
                result = False
        except Exception as error:
            result = False
        image = self.screen_shot(driver, "?????????????????????FAQ???????????????", result)
        self.log_infor("Result check page Contact: {}".format(result))
        self.log_infor("End check page Contact")
        return result, url, image

    def check_page_recruitment(self, driver, df):
        '''
        Check button Recruitment and return the result
        '''
        self.log_infor("Start check page Recruitment")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "???????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, 'r-rcp-header')
            title = driver.title
            url = driver.current_url
            if "????????????" in title:
                result = True
            else:
                result = False
        except Exception as error:
            result = False
        image = self.screen_shot(driver, "???????????????????????????", result)
        self.log_infor("Result check page Recruitment: {}".format(result))
        self.log_infor("End check page Recruitment")
        return result, url, image

    def check_button_english(self, driver, df):
        '''
        Check button English and return the result
        '''
        self.log_infor("Start check button English")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "English???????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            sleep(2)
            title = driver.title
            if "AOHATA" not in title:
                df_click = df[df["???????????????"].astype(str) == "English???????????????"]
                df_click.reset_index(inplace=True)
                self.click_by_df(driver, df_click)
                self.wait_page_complete(driver, 'c-hdg--lv1')
                title = driver.title
            url = driver.current_url
            str_company_summary = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div/div[2]/div/ul/li[1]/a').text
            str_our_business = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div/div[2]/div/ul/li[2]/a').text
            str_network = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div/div[2]/div/ul/li[3]/a').text
            str_products = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div/div[2]/div/ul/li[4]/a').text
            if ("AOHATA" in title) \
                    and (str_company_summary == "Company Summary")\
                    and (str_our_business == "Our Business")\
                    and (str_network == "Network & Related Companies")\
                    and (str_products == "Products"):
                result = True
            else:
                result = False
        except Exception as error:
            self.log_infor(error)
            result = False
        image = self.screen_shot(driver, "English???????????????", result)
        self.log_infor("Result check button English: {}".format(result))
        self.log_infor("End check button English")
        return result, url, image

    def check_button_chinese(self, driver, df):
        '''
        Check button Chinese and return the result
        '''
        self.log_infor("Start check button Chinese")
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "?????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, 'c-hdg--lv1')
            title = driver.title
            url = driver.current_url
            str_company_summary = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div/div[2]/div/ul/li[1]/a').text
            str_our_business = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div/div[2]/div/ul/li[2]/a').text
            str_network = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div/div[2]/div/ul/li[3]/a').text
            str_products = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div/div[2]/div/ul/li[4]/a').text
            if ("??????" in title) \
                    and (str_company_summary == "????????????")\
                    and (str_our_business == "????????????")\
                    and (str_network == "?????????????????????????????????????????????")\
                    and (str_products == "??????"):
                result = True
            else:
                result = False
        except Exception as error:
            self.log_infor(error)
            result = False
        image = self.screen_shot(driver, "?????????????????????", result)
        self.log_infor("Result check button Chinese: {}".format(result))
        self.log_infor("End check button Chinese")
        return result, url, image

    def check_button_next_slick_arrow(self, driver, df):
        '''
        Check button next slick arrow and return the result
        '''
        self.log_infor("Start check button next slick arrow")
        str_item_next = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "??????????????????????????????????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            sleep(3)
            str_item_next = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div[1]/div/div[1]/ul[2]/div/div/li[7]/a').text
            print(str_item_next)
            if '?????????????????????????????????' in str(str_item_next):
                result = True
            else:
                result = False
        except Exception as error:
            result = False
            self.log_infor(
                "Error on line " + str(sys.exc_info()[-1].tb_lineno) + " " + type(error).__name__ + " " + str(error)
            )
        image = self.screen_shot(driver, "??????????????????????????????????????????????????????", result)
        self.log_infor("Result check button next slick arrow: {}".format(result))
        self.log_infor("End check button next slick arrow")
        return result, str_item_next, image

    def check_button_prev_slick_arrow(self, driver, df):
        '''
        Check button next slick arrow and return the result
        '''
        self.log_infor("Start check button prev slick arrow")
        driver.get("https://www.aohata.co.jp/products/")
        str_item_prev = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "??????????????????????????????????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            sleep(2)
            str_item_prev = driver.find_element(By.XPATH, '//*[@id="bodyTop"]/div[2]/div[1]/div/div[1]/ul[2]/div/div/li[10]/a').text
            if "??????????????????????????????" in str(str_item_prev):
                result = True
            else:
                result = False
        except Exception as error:
            result = False
            self.log_infor(
                "Error on line " + str(sys.exc_info()[-1].tb_lineno) + " " + type(error).__name__ + " " + str(error)
            )
        image = self.screen_shot(driver, "??????????????????????????????????????????????????????", result)
        self.log_infor("Result check button prev slick arrow: {}".format(result))
        self.log_infor("End check button prev slick arrow")
        return result, str_item_prev, image

    def write_to_excel(self, output_file, col_value, col_report):
        '''
        write report to excel file with column report and column date
        col_report, col_date: Eg: I, K, ...
        '''
        from openpyxl import load_workbook
        work_book = load_workbook(output_file)
        work_sheet = work_book.active
        int_row = 3
        index = 0
        for result in self.results:
            # Write to col_report
            if str(result) == "True":
                work_sheet[col_report + str(int_row)].value = "OK"
            elif str(result) == "False":
                work_sheet[col_report + str(int_row)].value = "NG"
            else:
                work_sheet[col_report + str(int_row)].value = ""

            # Write to col_value
            if str(self.values[index]) == "":
                pass
            else:
                work_sheet[col_value + str(int_row)].value = str(self.values[index])
            # Write to col image_name and hyperlink image
            path_image = self.images_name[index]
            str_image_name = os.path.basename(self.images_name[index])
            work_sheet.cell(row=int_row, column=11).value = '=HYPERLINK("{}", "{}")'.format(path_image, str_image_name)     # column=11 is column K
            int_row += 1
            index += 1
        work_book.save(output_file)

    def check_first_prod(self, driver, df):
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "?????????????????????????????????????????????????????????????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, "l-pdc__intro")
            url = driver.current_url
            if url == "https://www.aohata.co.jp/products/55/":
                result = True
            else:
                result = False
        except:
            result = False
        image = self.screen_shot(driver, "?????????????????????????????????????????????????????????????????????????????????", result)
        return result, url, image

    def check_second_prod(self, driver, df):
        url = ""
        try:
            df_click = df[df["???????????????"].astype(str) == "????????????????????????????????????????????????????????????????????????????????????????????????"]
            df_click.reset_index(inplace=True)
            self.click_by_df(driver, df_click)
            self.wait_page_complete(driver, "l-pdc__intro")
            url = driver.current_url
            if url == "https://www.aohata.co.jp/products/marugoto/":
                result = True
            else:
                result = False
        except:
            result = False
        image = self.screen_shot(driver, "????????????????????????????????????????????????????????????????????????????????????????????????", result)
        return result, url, image

    def backup_and_create_output_file(self):
        '''
        move the output file into backup folder and create a new output file into the output folder
        '''
        self.log_infor("Start backup and create a new output file")
        output_file = self.curdir + "\\output\\output.xlsx"
        backup_directory = self.curdir + "\\bak\\output_" + datetime.datetime.now().strftime("%Y%m%d")
        self.create_folder(backup_directory)    # create a backup folder by day
        self.move_file(output_file, backup_directory)   # move the output file into the backup folder
        copyfile("D:\\TanLV\\SS1\\tmpl\\tmpl_scenario.xlsx", output_file)   # create a new output file by tmpl file
        self.log_infor("End backup and create a new output file")
        return output_file

    def click_by_df(self, driver, df):
        if str(df["ID"][0]) != "":
            driver.find_element(By.ID, str(df["ID"][0])).click()
        if str(df["CLASS"][0]) != "":
            driver.find_element(By.CLASS_NAME, str(df["CLASS"][0])).click()
        if str(df["TEXT"][0]) != "":
            driver.find_element(By.LINK_TEXT, str(df["TEXT"][0])).click()
        if str(df["XPATH"][0]) != "":
            driver.find_element(By.XPATH, str(df["XPATH"][0])).click()

    def get_attribute_by_df(self, driver, df, attribute):
        result = ""
        if str(df["???????????????"][0]) == "Check Banner 1":
            driver.switch_to.frame(driver.find_element(By.XPATH, '//*[@id="google_ads_iframe_/116070306/k/top_300x250_0"]'))
        if str(df["ID"][0]) != "":
            result = driver.find_element(By.ID, str(df["ID"][0])).get_attribute(attribute)
        elif str(df["CLASS"][0]) != "":
            result = driver.find_element(By.CLASS_NAME, str(df["CLASS"][0])).get_attribute(attribute)
        elif str(df["TEXT"][0]) != "":
            result = driver.find_element(By.LINK_TEXT, str(df["TEXT"][0])).get_attribute(attribute)
        elif str(df["XPATH"][0]) != "":
            result = driver.find_element(By.XPATH, str(df["XPATH"][0])).get_attribute(attribute)
        return result

    def check_banner(self, driver, df):
        folder_download = f"{self.curdir}\\download"
        path_image_original = f"{self.curdir}\\banner_image\\{str(df['???'][0])}"
        href_image = self.get_attribute_by_df(driver, df, "src")
        path_image_banner = self.download_image(href_image, folder_download)
        result = self.compare_images(path_image_original, path_image_banner)
        return result, path_image_banner

    def test(self):
        driver = self.open_browser_chrome()
        driver.maximize_window()
        driver.get("https://kakaku.com/")
        sleep(3)
        # driver.switch_to.frame(driver.find_element(By.XPATH, '//*[@id="google_ads_iframe_/116070306/k/top_300x250_0"]'))
        link = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[2]/div[4]/div[2]/div/p/img").get_attribute("src")
        print(link)

    def main_chrome(self, output_file):
        df = pd.read_excel(output_file, sheet_name=0, header=0, dtype=object)
        df = df.replace(np.nan,"")
        driver = self.open_browser_chrome()
        driver.maximize_window()
        for id, row in df.iterrows():
            result = ""
            value = ""
            image_name = ""
            if str(row["URL"]) != "":
                url = row["URL"]
            driver.get(url)
            sleep(2)
            if str(row["???????????????"]) == "???????????????????????????":
                result, value, image_name = self.check_page_recommend_products(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "????????????????????????????????????":
                result, value, image_name = self.check_page_recommend_recipe(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "?????????????????????????????????????????????":
                result, value, image_name = self.check_page_experience(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "???????????????????????????":
                result, value, image_name = self.check_page_company(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "?????????????????????FAQ???????????????":
                result, value, image_name = self.check_page_contact(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "???????????????????????????":
                result, value, image_name = self.check_page_recruitment(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "English???????????????":
                result, value, image_name = self.check_button_english(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "?????????????????????":
                result, value, image_name = self.check_button_chinese(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "???????????????????????????????????????????????????":
                result, value, image_name = self.check_logo_home_page_above(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "??????????????????????????????????????????????????????":
                result, value, image_name = self.check_button_next_slick_arrow(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "??????????????????????????????????????????????????????":
                driver.get("https://www.aohata.co.jp/")
                result, value, image_name = self.check_button_prev_slick_arrow(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "?????????????????????????????????????????????????????????????????????????????????":
                result, value, image_name = self.check_first_prod(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "????????????????????????????????????????????????????????????????????????????????????????????????":
                result, value, image_name = self.check_second_prod(driver, df)
                sleep(2)
            if str(row["???????????????"]) == "Check Banner 1":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 1"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1
            if str(row["???????????????"]) == "Check Banner 2":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 2"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1
            if str(row["???????????????"]) == "Check Banner 3":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 3"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1
            if str(row["???????????????"]) == "Check Banner 4":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 4"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1
            if str(row["???????????????"]) == "Check Banner 5":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 5"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1
            if str(row["???????????????"]) == "Check Banner 6":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 6"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1
            if str(row["???????????????"]) == "Check Banner 7":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 7"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1
            if str(row["???????????????"]) == "Check Banner 8":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 8"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1
            if str(row["???????????????"]) == "Check Banner 9":
                df_click = df[df["???????????????"].astype(str) == "Check Banner 9"]
                df_click.reset_index(inplace=True)
                result, path_image = self.check_banner(driver, df_click)
                image_name = f"{self.curdir}\\image\\{self.int_count}-{str(df_click['???????????????'][0])}-{str(result)}.png"
                copyfile(path_image, image_name)
                self.int_count += 1

            self.values.append(value)
            self.results.append(result)
            self.images_name.append(image_name)
        self.write_to_excel(output_file, "I", "J")
        driver.quit()

if __name__ == "__main__":
    curdir = os.path.dirname(
            os.path.abspath(inspect.getfile(inspect.currentframe())))
    driver = Browser(curdir, LOG_INFO)
    output_file = driver.backup_and_create_output_file()
    driver.main_chrome(output_file)
    # driver.test()