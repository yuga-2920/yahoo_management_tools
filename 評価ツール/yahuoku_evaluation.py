from datetime import datetime
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import os
import re
import eel
import traceback
import time

from line_setting import line_error_notification
from format_spreadsheet import set_main_spreadsheet, df_main_spreadsheet, write_main_spreadsheet, get_main_spreadsheet

COMMENT = "ありがとうございました。とても良い取引ができました。また機会がありましたら、よろしくお願いいたします。"

#Cromeの設定
def set_driver(driver_path):

    # Chromeドライバーの読み込み
    options = ChromeOptions()

    # 起動オプションの設定
    # options.add_argument('--headless')
    # options.add_argument('--single-process')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('log-level=3')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--incognito')
    options.add_argument('--user-agent=Chrome/87.0.42word_num.88')
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-web-security')
    options.add_argument('--disable-desktop-notifications')
    options.add_argument('--disable-application-cache')
    options.add_argument("--disable-extensions")
    options.add_argument('--lang=ja')

    # ChromeのWebDriverオブジェクトを作成する。
    return Chrome(driver_path, options=options)

#urlの整形
def  format_url(url):
    url_replace = url.replace('https://contact.auctions.yahoo.co.jp/seller/top?aid=', '')
    user = re.sub(r'([\w]+)&syid=', '', url_replace)
    user = re.sub(r'&bid=([\w-]+)', '', user)
    user = re.sub(r'&oid=([\w-]+)', '', user)
    if user == 'vvble04942':
        password = 'Masashi09120214'
    elif user == 'aazrl67349':
        password = 'Masashi0912!'

    return user, password

#ログイン関数
def login(driver, wait, user, password):
    time.sleep(3)

    #ID/携帯電話/メールアドレスの入力
    wait.until(EC.visibility_of_element_located((By.ID, 'username')))
    username = driver.find_element_by_id("username")
    username.send_keys(user)

    wait.until(EC.visibility_of_element_located((By.ID, 'btnNext')))
    driver.find_element_by_id("btnNext").click()

    #パスワードの入力
    wait.until(EC.visibility_of_element_located((By.ID, 'passwd')))
    passwd = driver.find_element_by_id("passwd")
    passwd.send_keys(password)

    wait.until(EC.visibility_of_element_located((By.ID, 'loginSubmit')))
    loginSubmit = driver.find_element_by_id("loginSubmit")

    wait.until(EC.visibility_of_element_located((By.ID, 'btnSubmit')))
    btnSubmit = loginSubmit.find_element_by_id("btnSubmit")
    btnSubmit.click()

#発送連絡
def shipping_contact(driver, wait):

    try:
        driver.execute_script('document.querySelector("#plibLoadMdlInner > div > div > input").click()')

    except:
        pass

    #商品を発送しましたボタンのクリック
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    wait.until(EC.visibility_of_element_located((By.ID, 'jsCheckSend')))
    jsCheckSend = driver.find_element_by_id('jsCheckSend')
    jsCheckSend.click()
    time.sleep(1)

    #発送連絡ボタンのクリック
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'jsOnSendButton')))
    jsOnSendButton = driver.find_element_by_class_name('jsOnSendButton')
    jsOnSendButton.click() 
    time.sleep(1)

#評価
def evaluation(driver, wait):

    try:
        driver.execute_script('document.querySelector("#plibLoadMdlInner > div > div > input").click()')

    except:
        pass

    #落札者評価ボタンのクリック
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'libBtnBlueL')))
    libBtnBlueL = driver.find_element_by_class_name("libBtnBlueL")
    
    if libBtnBlueL.text == "落札者を評価する":
        libBtnBlueL.click()
        time.sleep(1)

    #評価
    wait.until(EC.visibility_of_element_located((By.ID, 'rating')))
    options = driver.find_element_by_id("rating")
    select = Select(options)
    select.select_by_value("veryGood")
    time.sleep(1)

    #コメント
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    wait.until(EC.visibility_of_element_located((By.NAME, 'formComment')))
    formComment = driver.find_element_by_name("formComment")
    formComment.send_keys(COMMENT)
    time.sleep(1)

    #確認ボタンのクリック
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'libBtnBlueL')))
    libBtnBlueL = driver.find_element_by_class_name("libBtnBlueL")
    libBtnBlueL.click()
    time.sleep(1)

    #評価ボタンの公開
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'libBtnBlueL')))
    libBtnBlueL = driver.find_element_by_class_name("libBtnBlueL")
    libBtnBlueL.click()
    time.sleep(1)

#発送連絡と評価
def ship_contact_and_evaluation(driver, wait, worksheet, url, user, password, index, i):

    #urlにアクセス
    driver.get(url)

    #ログイン
    login(driver, wait, user, password)

    try:
        #発送連絡
        shipping_contact(driver, wait)
        #評価
        evaluation(driver, wait)

    except TimeoutException:

        try:
            #評価
            evaluation(driver, wait)

        except TimeoutException:
            auction_bool = driver.find_element_by_class_name("decJS").text
            if auction_bool == "このオークションはまだ終了していません。":
                #ログの表示
                now = datetime.now()
                write_main_spreadsheet(worksheet, "R{}".format(index + 2), now.strftime("%Y/%m/%d %H:%M:%S"))
                
                eel.view_log_js("{}枚目の商品はオークション中のため評価できませんでした\n".format(i + 1))
                return False

            else:
                eel.view_log_js("{}枚目の商品は評価できませんでした\n".format(i + 1))
                t = traceback.format_exc()
                line_error_notification('発送連絡評価ツール\n' + t + '\n"予期せぬエラーが発生しました。"')
                return False

    #書き込む日付の取得
    now = datetime.now()

    #スプレッドシートに書き込み
    value = get_main_spreadsheet(worksheet, "R{}".format(index + 2))
    if value is None:
        write_main_spreadsheet(worksheet, "R{}".format(index + 2), now.strftime("%Y/%m/%d %H:%M:%S"))
        write_main_spreadsheet(worksheet, "AA{}".format(index + 2), now.strftime("%Y/%m/%d %H:%M:%S"))

    else:
        write_main_spreadsheet(worksheet, "AA{}".format(index + 2), now.strftime("%Y/%m/%d %H:%M:%S"))

    return True

def main():

    try:
        #評価するメールIDの読み込み
        eval_item_list = eel.col_list()().split('\n')

        #ワークシートの取得
        worksheet = set_main_spreadsheet()

        #スプレッドシートのデータの読み込み
        df = df_main_spreadsheet(worksheet)

        #driverの設定
        driver_path = ChromeDriverManager().install()

        for i, item_info in enumerate(eval_item_list):

            #ドライバーの設定
            driver = set_driver(driver_path)
            driver.implicitly_wait(5)
            driver.set_script_timeout(5)
            wait = WebDriverWait(driver, 3)

            #各商品の情報
            item_info_list = item_info.split(',')
            order_no = item_info_list[0]
            order_pass = item_info_list[1]

            try:
                #urlの取得
                df1 = df[(df['Ship_Order_No'] == order_no)]
                item_col = df1[(df1['Ship_Order_Pass'] == order_pass)]
                index = item_col.index[0]
                url = item_col['Navi'].iloc[0]
            
            except IndexError:
                eel.view_log_js("{}枚目の商品はありません。\n".format(i + 1))
                continue

            #アクセスユーザー名とパスワードの取得
            user, password = format_url(url)

            try:
                #ツールの実行
                bool = ship_contact_and_evaluation(driver, wait, worksheet, url, user, password, index,i)

                if bool:
                    #ログの表示
                    eel.view_log_js("{}枚目の商品を評価しました\n".format(i + 1))

                # driver.quit()

            except:
                eel.view_log_js("{}枚目の商品は評価できませんでした\n".format(i + 1))
                t = traceback.format_exc()
                line_error_notification('発送連絡評価ツール')
                t_list = [t[i: i + 1000] for i in range(0, len(t), 1000)]
                for message in t_list:
                    line_error_notification(message)
                line_error_notification('予期せぬエラーが発生しました。')

        else:
            eel.view_log_js("終了しました")

    except:
        t = traceback.format_exc()
        line_error_notification('発送連絡評価ツール\n' + t + '\n"予期せぬエラーが発生しました。"')