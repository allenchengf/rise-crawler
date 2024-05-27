from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import os, stat
import requests
import time
import logging
import platform
import csv


def main():
    correct_path = get_path_by_os()
    current_directory = os.getcwd() + correct_path + "temp"
    chrome_options = Options()

    # 不需開啟瀏覽器
    # chrome_options.add_argument('--headless=new')
    # chrome_options.add_argument('--no-sandbox')

    chrome_options.add_argument('--disable-extensions')
    prefs = {"download.default_directory": current_directory}
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')

    s = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=s, options=chrome_options)
    driver.maximize_window()

    # Send a get request to the url
    driver.get('https://rise.iii.org.tw/rise_front/#/view/form/z7rxvsrewmd6s3r2')
    original_window = driver.current_window_handle
    driver.implicitly_wait(90)
    # time.sleep(3)

    logging_message("Start Process:")
    login = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/button[1]')
    driver.execute_script("arguments[0].click();", login)

    guestLogin = driver.find_element(By.XPATH, '//*[@id="all_2"]/div/div[2]/div[3]/div/button[1]')
    driver.execute_script("arguments[0].click();", guestLogin)

    time.sleep(2)
    menuButton = driver.find_element(By.XPATH, '//*[@id="all_2"]/div[1]/div[10]/table/tbody/td/p/a')
    driver.execute_script("arguments[0].click();", menuButton)

    with open("store.csv", "r", encoding="utf-8") as csvfile:
        csvreader = csv.reader(csvfile)
        next(csvreader)
        next(csvreader)
        for index, row in enumerate(csvreader):
            driver.switch_to.new_window('tab')
            driver.get('https://rise.iii.org.tw/rise_front/#/view/form/btip1i8j523d99uu')

            try:
                # 是否同意
                time.sleep(2)
                agree = driver.find_element(By.XPATH, '//*[@id="read-first"]/div[2]/div[1]/input')
                driver.execute_script("arguments[0].click();", agree)
                start = driver.find_element(By.XPATH, '// *[ @ id = "read-first"] / div[2] / div[2] / button')
                driver.execute_script("arguments[0].click();", start)

                # 開始作答
                time.sleep(2)
                logging_message("Input " + str(index) + " row data")
                logging_message(row)
                # 1. 請問，貴公司名稱
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[2]/div/div/div/div[2]/div/input').send_keys(
                    row[11])
                # 2. 請問，貴公司統一編號？
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[3]/div/div/div/div[2]/div/input').send_keys(
                    row[12])

                # 3. 請問，目前公司銷售或服務的對象？
                q3 = driver.find_element(By.XPATH,
                                         '//*[@id="app"]/div/div[3]/div[2]/div[4]/div/div/div/div[2]/div/div[' + row[
                                             13] + ']/label/input')
                driver.execute_script("arguments[0].click();", q3)

                # 4. 請問，目前公司經常雇用員工人數(有勞健保)？
                q4 = driver.find_element(By.XPATH,
                                         '//*[@id="app"]/div/div[3]/div[2]/div[5]/div/div/div/div[2]/div/div[' + row[
                                             14] + ']/label/input')
                driver.execute_script("arguments[0].click();", q4)

                # 5. 請問，公司是屬於下列哪一種產業別？
                q5 = driver.find_element(By.XPATH,
                                         '//*[@id="app"]/div/div[3]/div[2]/div[6]/div/div/div/div[2]/div/div['
                                         '16]/label/input')
                driver.execute_script("arguments[0].click();", q5)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[6]/div/div/div/div[2]/div/div['
                                    '16]/label/div/input').send_keys('電子通訊／電腦週邊零售業')

                # 6. 請問，貴公司去年營收狀況?
                q6 = driver.find_element(By.XPATH,
                                         '//*[@id="app"]/div/div[3]/div[2]/div[7]/div/div/div/div[2]/div/div[' + row[
                                             16] + ']/label/input')
                driver.execute_script("arguments[0].click();", q6)

                # 7. 請問，貴公司目前網實銷售通路營收占比狀況？ 【請於空格填入0-100，請注意三個數字加總需為100%】
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/div[1]/label/input').send_keys(
                    row[17])
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/div[2]/label/input').send_keys(
                    row[18])
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/div[3]/label/input').send_keys(
                    row[19])

                # 換頁
                nextPage = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div[2]/div[9]/button')
                driver.execute_script("arguments[0].click();", nextPage)

                # 2_1. 與同業相比較，貴公司在「作業管理及研發方面」數位化採用的程度如何？ 請以0-5分評比，0分代表大幅落後同業，5分代表大幅領先同業。 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[2]/div/div/div/div[2]/div/input').send_keys(
                    row[20])

                # 2_2. 與同業相比較，貴公司在「行銷與銷售方面」數位化採用的程度如何？ 請以0-5分評比，0分代表大幅落後同業，5分代表大幅領先同業。 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[3]/div/div/div/div[2]/div/input').send_keys(
                    row[21])

                # 2_3. 與同業相比較，貴公司在「人事管理方面」數位化採用的程度如何？ 請以0-2分評比，0分代表大幅落後同業，2分代表大幅領先同業。 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[4]/div/div/div/div[2]/div/input').send_keys(
                    row[22])

                # 2_4. 與同業相比較，貴公司在「財務管理方面」數位化採用的程度如何？ 請以0-3分評比，0分代表大幅落後同業，3分代表大幅領先同業。 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[5]/div/div/div/div[2]/div/input').send_keys(
                    row[23])

                # 2_5. 與同業相比較，貴公司在「決策數據分析方面」數位化採用的程度如何？ 請以0-3分評比，0分代表大幅落後同業，3分代表大幅領先同業。 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[6]/div/div/div/div[2]/div/input').send_keys(
                    row[24])

                # 2_6. 與同業相比較，貴公司在「決策數據分析方面」數位化採用的程度如何？ 請以0-3分評比，0分代表大幅落後同業，3分代表大幅領先同業。 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[7]/div/div/div/div[2]/div/input').send_keys(
                    row[25])

                # 2_7. 請問，貴公司在顧客服務/顧客維繫/銷售/行銷等，有採用下列哪些軟體或工具？
                q2_7_1 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/div['
                                             '1]/label/input')
                driver.execute_script("arguments[0].click();", q2_7_1)

                q2_7_2 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/div['
                                             '2]/label/input')
                driver.execute_script("arguments[0].click();", q2_7_2)
                q2_7_3 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/div['
                                             '3]/label/input')
                driver.execute_script("arguments[0].click();", q2_7_3)
                q2_7_4 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/div['
                                             '4]/label/input')
                driver.execute_script("arguments[0].click();", q2_7_4)
                q2_7_5 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/div['
                                             '5]/label/input')
                driver.execute_script("arguments[0].click();", q2_7_5)

                # 2_8. 請問，貴公司在現場作業管理、供應鏈管理、研發等，有採用下列哪些軟體或工具？
                q2_8_1 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[9]/div/div/div/div[2]/div/div['
                                             '2]/label/input')
                driver.execute_script("arguments[0].click();", q2_8_1)

                q2_8_2 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[9]/div/div/div/div[2]/div/div['
                                             '3]/label/input')
                driver.execute_script("arguments[0].click();", q2_8_2)
                q2_8_3 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[9]/div/div/div/div[2]/div/div['
                                             '4]/label/input')
                driver.execute_script("arguments[0].click();", q2_8_3)

                # 2_9. 請問，貴公司在人事管理/人力資源管理等，有採用下列哪些軟體或工具？
                q2_9_1 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[10]/div/div/div/div[2]/div/div['
                                             '1]/label/input')
                driver.execute_script("arguments[0].click();", q2_9_1)

                q2_9_2 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[10]/div/div/div/div[2]/div/div['
                                             '2]/label/input')
                driver.execute_script("arguments[0].click();", q2_9_2)
                q2_9_3 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[10]/div/div/div/div[2]/div/div['
                                             '3]/label/input')
                driver.execute_script("arguments[0].click();", q2_9_3)

                # 2_10. 請問，貴公司在財務管理等，有採用下列哪些軟體或工具？
                q2_10_1 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[11]/div/div/div/div[2]/div/div['
                                              '2]/label/input')
                driver.execute_script("arguments[0].click();", q2_10_1)

                q2_10_2 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[11]/div/div/div/div[2]/div/div['
                                              '3]/label/input')
                driver.execute_script("arguments[0].click();", q2_10_2)
                q2_10_3 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[11]/div/div/div/div[2]/div/div['
                                              '4]/label/input')
                driver.execute_script("arguments[0].click();", q2_10_3)

                # 2_11. 請問，貴公司在財務管理等，有採用下列哪些軟體或工具？
                q2_11_1 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[12]/div/div/div/div[2]/div/div['
                                              '1]/label/input')
                driver.execute_script("arguments[0].click();", q2_11_1)

                q2_11_2 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[12]/div/div/div/div[2]/div/div['
                                              '2]/label/input')
                driver.execute_script("arguments[0].click();", q2_11_2)
                q2_11_3 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[12]/div/div/div/div[2]/div/div['
                                              '3]/label/input')
                driver.execute_script("arguments[0].click();", q2_11_3)

                # 2_12. 請問，貴公司是否有將上述蒐集來的數據做哪些分析與運用？
                q2_12 = driver.find_element(By.XPATH,
                                            '//*[@id="app"]/div/div[3]/div[2]/div[13]/div/div/div/div[2]/div/div[' +
                                            row[31] + ']/label/input')
                driver.execute_script("arguments[0].click();", q2_12)

                # 2_13. 請問，貴公司在內部溝通方面，有運用哪些數位工具或軟體？
                q2_13_1 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[14]/div/div/div/div[2]/div/div['
                                              '1]/label/input')
                driver.execute_script("arguments[0].click();", q2_13_1)

                q2_13_2 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[14]/div/div/div/div[2]/div/div['
                                              '2]/label/input')
                driver.execute_script("arguments[0].click();", q2_13_2)

                q2_13_3 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[14]/div/div/div/div[2]/div/div['
                                              '3]/label/input')
                driver.execute_script("arguments[0].click();", q2_13_3)

                q2_13_4 = driver.find_element(By.XPATH,
                                              '//*[@id="app"]/div/div[3]/div[2]/div[14]/div/div/div/div[2]/div/div['
                                              '4]/label/input')
                driver.execute_script("arguments[0].click();", q2_13_4)

                # 2_14. 請問，貴公司預計今年投入在數位工具上的預算，與去年相比？
                q2_14 = driver.find_element(By.XPATH,
                                            '//*[@id="app"]/div/div[3]/div[2]/div[15]/div/div/div/div[2]/div/div[' +
                                            row[33] + ']/label/input')
                driver.execute_script("arguments[0].click();", q2_14)

                # 換頁
                nextPage = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div[2]/div[16]/button[2]')
                driver.execute_script("arguments[0].click();", nextPage)

                # 3_1. 請問，貴公司在資訊方面的人力配置狀況？
                q3_1 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[2]/div/div/div/div[2]/div/div[' + row[
                                               34] + ']/label/input')
                driver.execute_script("arguments[0].click();", q3_1)

                # 3_2. 請問，貴公司在資訊方面的人力配置狀況？
                q3_2 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[3]/div/div/div/div[2]/div/div[' + row[
                                               35] + ']/label/input')
                driver.execute_script("arguments[0].click();", q3_2)

                # 3_3. 請問，貴公司在資訊方面的人力配置狀況？
                q3_3 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[4]/div/div/div/div[2]/div/div[' + row[
                                               36] + ']/label/input')
                driver.execute_script("arguments[0].click();", q3_3)

                # 3_4. 請問，貴公司未來有意願培育下列哪些領域的數位人才？
                q3_4_1 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[5]/div/div/div/div[2]/div/div['
                                             '1]/label/input')
                driver.execute_script("arguments[0].click();", q3_4_1)

                q3_4_2 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[5]/div/div/div/div[2]/div/div['
                                             '2]/label/input')
                driver.execute_script("arguments[0].click();", q3_4_2)
                q3_4_3 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[5]/div/div/div/div[2]/div/div['
                                             '3]/label/input')
                driver.execute_script("arguments[0].click();", q3_4_3)

                # 換頁
                nextPage = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div[2]/div[6]/button[2]')
                driver.execute_script("arguments[0].click();", nextPage)

                # 4_1. 貴公司導入之行動應用服務系統介面(如系統操作頁面)是簡單、容易操作的。
                q4_1 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[3]/div/div/div/div[2]/div/div[' + row[
                                               38] + ']/label/input')
                driver.execute_script("arguments[0].click();", q4_1)

                # 4_2. 貴公司所導入之行動應用服務，整體來說是實用、有幫助的。
                q4_2 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[4]/div/div/div/div[2]/div/div[' + row[
                                               39] + ']/label/input')
                driver.execute_script("arguments[0].click();", q4_2)

                # 4_3. 貴公司對目前使用的行動應用服務整體感到滿意。
                q4_3 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[5]/div/div/div/div[2]/div/div[' + row[
                                               40] + ']/label/input')
                driver.execute_script("arguments[0].click();", q4_3)

                # 4_4. 請問，導入本次行動應用服務後，貴公司的營運成本降低比例？ 【請填入0-100】 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[7]/div/div/div/div[2]/div/input').send_keys(
                    row[41])

                # 4_5. 請問，導入本次行動應用服務後，貴公司的來客數增加比例？ 【請填入0-100】 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[8]/div/div/div/div[2]/div/input').send_keys(
                    row[42])

                # 4_6. 請問，導入本次行動應用服務後，貴公司的營業額增加比例？ 【請填入0-100】 (請輸入整數)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[9]/div/div/div/div[2]/div/input').send_keys(
                    row[43])

                # 4_7. 請問，貴公司未來會持續使用本次導入的行動應用服務嗎？
                q4_7 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[11]/div/div/div/div[2]/div/div[' + row[
                                               44] + ']/label/input')
                driver.execute_script("arguments[0].click();", q4_7)

                # 4_8. 請問，減少貴公司使用此服務意願的原因為何？
                q4_8 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[12]/div/div/div/div[2]/div/div[' + row[
                                               45] + ']/label/input')
                driver.execute_script("arguments[0].click();", q4_8)

                # 4_9. 請問，貴公司希望未來可以有哪些行動智慧應用服務？
                q4_9_1 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[13]/div/div/div/div[2]/div/div[1]/label/input')
                driver.execute_script("arguments[0].click();", q4_9_1)

                q4_9_2 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[13]/div/div/div/div[2]/div/div[2]/label/input')
                driver.execute_script("arguments[0].click();", q4_9_2)

                q4_9_3 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[13]/div/div/div/div[2]/div/div[3]/label/input')
                driver.execute_script("arguments[0].click();", q4_9_3)

                q4_9_4 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[13]/div/div/div/div[2]/div/div[4]/label/input')
                driver.execute_script("arguments[0].click();", q4_9_4)

                q4_9_5 = driver.find_element(By.XPATH,
                                             '//*[@id="app"]/div/div[3]/div[2]/div[13]/div/div/div/div[2]/div/div[5]/label/input')
                driver.execute_script("arguments[0].click();", q4_9_5)

                # 換頁
                nextPage = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div[2]/div[14]/button[2]')
                driver.execute_script("arguments[0].click();", nextPage)

                # 5_1. 請問填表人姓名？
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[2]/div/div/div/div[2]/div/input').send_keys(
                    row[47])

                # 5_2. 請問填表人聯絡EMAIL？(如：XXX@XXX.com、XXX@XXX.com.XX)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[3]/div/div/div/div[2]/div/input').send_keys(
                    row[48])

                # 5_3. 請問填表人聯絡電話號碼？(手機格式：0910123123；室內電話格式：02-12341234#123)
                driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div/div[3]/div[2]/div[4]/div/div/div/div[2]/div/input').send_keys(
                    row[49])

                # 5_4. 貴公司所在縣市
                q5_4 = driver.find_element(By.XPATH,
                                           '//*[@id="app"]/div/div[3]/div[2]/div[5]/div/div/div/div[2]/div/div['+row[50]+']/label/input')
                driver.execute_script("arguments[0].click();", q5_4)

                # 送出
                submit = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div[2]/div[7]/button')
                driver.execute_script("arguments[0].click();", submit)
            except requests.exceptions.RequestException as e:
                logging_message(e)
                print(e)

        # time.sleep(30)
        # driver.quit()
        print("Process Done!")


def get_path_by_os():
    os_name = platform.system()
    if os_name == 'Windows':
        return '\\'
    else:
        return '/'


def logging_message(message):
    print(message)
    logging.basicConfig(level=logging.INFO, filename='accesslog ' + time.strftime('%Y%m%d_%H_%M_%S') + '.log',
                        filemode='a', format='%(asctime)s %(levelname)s: %(message)s')
    logging.info(message)


if __name__ == '__main__':
    main()
