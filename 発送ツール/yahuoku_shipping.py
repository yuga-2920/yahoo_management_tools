from datetime import datetime
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import os
import re
import traceback
import time
import sys
import chromedriver_binary
import eel

from line_setting import line_notification, line_error_notification
from format_spreadsheet import df_main_spreadsheet, write_main_spreadsheet
from setting_pdf import write_single_item_pdf, write_group_item_pdf


#ドライバーの設定
def set_driver(driver_path):

    # Chromeドライバーの読み込み
    options = ChromeOptions()

    # 起動オプションの設定
    # options.add_argument('--headless')
    # options.add_argument('--single-process')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('log-level=3')
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--incognito')
    options.add_argument('--user-agent=Chrome/87.0.42word_num.88')
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
def format_url(url):
    url_replace = url.replace('https://contact.auctions.yahoo.co.jp/seller/top?aid=', '')
    user = re.sub(r'([\w]+)&syid=', '', url_replace)
    user = re.sub(r'&bid=([\w]+)', '', user)
    user = re.sub(r'&oid=([\w-]+)', '', user)
    if user == 'vvble04942':
        password = 'Masashi09120214'
    elif user == 'aazrl67349':
        password = 'Masashi0912!'

    return user, password

#商品名の処理
def format_item_name(item_name):
    item_name_list = [item_name[i: i + 20] for i in range(0, len(item_name), 20)]

    if len(item_name_list) != 3:
        difference = 3 - len(item_name_list)
        for _ in range(difference):
            item_name_list = [""] + item_name_list

    return item_name_list[::-1]

#終了日時をスプレッドシートから取得できる形式に変更
def change_date_fetche_spreadsheet(date):
    date = date.replace("終了日時：", "").replace(" ", "")
    date = datetime.strptime(date, "%m月%d日%H時%M分")
    return date.strftime("%m/%d %H:%M")

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

# 商品の発送処理
def before_create_pdf_process(driver, wait, url):

    #ポップアップの削除
    try:
        driver.execute_script('document.querySelector("#plibLoadMdlInner > div > div.alignC > p:nth-child(2) > input").click()')
    except:
        pass

    try:

        #品名の入力
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'decItmName')))
        item_name = driver.find_element_by_class_name("decItmName").text

        if "●" in item_name:
            item_name_split = item_name.split("●")[-1]
            wait.until(EC.visibility_of_element_located((By.ID, 'inputItemName')))
            inputItemName = driver.find_element_by_id("inputItemName")
            inputItemName.send_keys(item_name_split)

        else:
            #出品できなかった旨をLINEに通知
            line_notification('商品名に●が記載されていないため\nURL:{}\nは発送コードを発行できませんでした'.format(url))
            return

        #配送コードの発行
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'libBtnBlueL')))
        libBtnBlueL = driver.find_element_by_class_name("libBtnBlueL")
        libBtnBlueL.click()
    
    except TimeoutException:
        pass

    try:
        #ポップアップの削除
        driver.execute_script('document.querySelector("#redSubmitBtn").click()')
    except:
        pass

#スクリーンショット
def screenshot(driver, save_name):
    driver.execute_script('window.scroll(0,500);') 
    driver.save_screenshot('スクリーンショット\\{}.png'.format(save_name))

    html = driver.page_source
    with open("HTML\\{}.html".format(save_name), "w", encoding='utf-8') as f:
        f.write(html)

    img = driver.find_element_by_id("yamatoQr").screenshot_as_png
    with open("QRcode\\{}.png".format(save_name), "wb") as f:
        f.write(img)

#単品商品のpdf作成とスプレッドシートに記入
def single_item_pdf_spreadsheet(df, mail_id, item_name_list, delivery_method, delivery_time, receipt_num, item_password, src):
    

    row = df[df['Mail_ID'] == mail_id]
    index = row.index[0]
    mail_id = row['Mail_ID'].values[0]
    product_id = row['Product_ID'].values[0]
    quantity = row['Quantity'].values[0]
    order_no = row['Order_No'].values[0]

    item_description_list = [product_id, quantity, order_no]

    #書き込む日付の取得
    now = datetime.now()

    #スプレッドシートに書き込み
    write_main_spreadsheet("X{}".format(index + 2), item_password)
    write_main_spreadsheet("W{}".format(index + 2), receipt_num)
    write_main_spreadsheet("Y{}".format(index + 2), src)
    write_main_spreadsheet("Z{}".format(index + 2), now.strftime("%Y/%m/%d %H:%M:%S"))
    write_main_spreadsheet("AB{}".format(index + 2), delivery_time)
    write_main_spreadsheet("AC{}".format(index + 2), delivery_method)

    #pdfを作成
    write_single_item_pdf(mail_id, item_name_list, delivery_method, delivery_time, receipt_num, item_password, item_description_list)

#まとめて取引のpdf作成とスプレッドシートに記入
def multiple_items_pdf_spreadsheet(driver, wait, origin_df, df, mail_id, delivery_method, delivery_time, receipt_num, item_password, src):

    row = df[df['Mail_ID'] == mail_id]
    index = row.index[0]

    #まとめ取り引きの商品の確認
    rapidnofollow = driver.find_element_by_class_name("rapidnofollow")
    url = rapidnofollow.get_attribute("href")

    driver.get(url)

    #何をしているのか確認
    while True:
        try:
            libJsExpandToggleBtn = driver.find_element_by_class_name('libJsExpandToggleBtn')
            a = libJsExpandToggleBtn.find_element_by_tag_name('a')
            a.click()
        except NoSuchElementException:
            break

    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'lxml')

    #落札者の取得
    decBuyerID = soup.find(class_ = "decBuyerID")
    purchaser = decBuyerID.get_text(strip=True).replace("落札者：", "").replace(" ", "").split('（')[0]

    decBundle = soup.find(class_ = "decBundle")
    item_list = decBundle.find_all('tr')

    items_description_list = []
    for i, item in enumerate(item_list):
        item_description_list = []

        #商品名
        decItmName = item.find(class_ = "decItmName")
        item_name = decItmName.get_text(strip=True)
        item_name_list = format_item_name(item_name)
        item_description_list.append(item_name_list)

        #商品の行を取得するための操作
        href = decItmName.find("a")['href']
        # end_day = item.find(class_ = "decMDT").get_text(strip=True)
        # end_day = change_date_fetche_spreadsheet(end_day)

        row = origin_df.query('Reference.str.contains("{}") & Purchaser.str.contains("{}")'.format(href, purchaser))
        index = row.index[0]

        #pdf作成に必要な情報の取得
        product_id = row['Product_ID'].values[0]
        quantity = row['Quantity'].values[0]
        order_no = row['Order_No'].values[0]

        item_description_list.append(product_id)
        item_description_list.append(quantity)
        item_description_list.append(order_no)

        items_description_list.append(item_description_list)

        if i == 0:
            #書き込む日付の取得
            now = datetime.now()

            #スプレッドシートに書き込み
            write_main_spreadsheet("W{}".format(index + 2), receipt_num)
            write_main_spreadsheet("X{}".format(index + 2), item_password)
            write_main_spreadsheet("Y{}".format(index + 2), src)
            write_main_spreadsheet("Z{}".format(index + 2), now.strftime("%Y/%m/%d %H:%M:%S"))
            write_main_spreadsheet("AB{}".format(index + 2), delivery_time)
            write_main_spreadsheet("AC{}".format(index + 2), delivery_method)

        else:
            #スプレッドシートに書き込み
            write_main_spreadsheet("Q{}".format(index + 2), "まとめて")
            write_main_spreadsheet("R{}".format(index + 2), "まとめて")
            write_main_spreadsheet("S{}".format(index + 2), "まとめて")
            write_main_spreadsheet("W{}".format(index + 2), "まとめて")
            write_main_spreadsheet("X{}".format(index + 2), "まとめて")
            write_main_spreadsheet("Y{}".format(index + 2), "まとめて")
            write_main_spreadsheet("Z{}".format(index + 2), "まとめて")
            write_main_spreadsheet("AA{}".format(index + 2), "まとめて")
            write_main_spreadsheet("AB{}".format(index + 2), "まとめて")
            write_main_spreadsheet("AC{}".format(index + 2), "まとめて")

    #pdfを作成
    write_group_item_pdf(mail_id, delivery_method, delivery_time, receipt_num, item_password, items_description_list)

#pdf作成とスプレッドシートに書き込み
def create_pdf_write_spreadsheet(driver, wait, origin_df, df, row):

    mail_id = row['Mail_ID']

    try:
        driver.execute_script('document.querySelector("#plibLoadMdlInner > div > div > input").click()')
    except:
        pass

    #発送場所をヤマト運輸にする
    wait.until(EC.visibility_of_element_located((By.ID, 'codeSelect')))
    options = driver.find_element_by_id("codeSelect")
    select = Select(options)
    select.select_by_value("yamato")

    #pdfとスプレッドシートに書き込む情報の取得
    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'lxml')

    src = soup.find(id = 'yamatoQr')['src']
    shipSelect__right = soup.find(class_ = "shipSelect__right")
    table = shipSelect__right.find("table")
    tds = table.find_all("td")

    if len(tds) == 5:
        delivery_method = tds[0].get_text(strip=True).replace("配送方法を変更", "")
        delivery_time = tds[1].get_text(strip=True)
        receipt_num = tds[2].get_text(strip=True)
        item_password = tds[3].get_text(strip=True)

        if len(receipt_num) == 9:
            receipt_num = "0" + str(receipt_num)

    if len(tds) == 4:
        delivery_method = tds[0].get_text(strip=True).replace("配送方法を変更", "")
        delivery_time = "記載なし"
        receipt_num = tds[1].get_text(strip=True)
        item_password = tds[2].get_text(strip=True)

        if len(receipt_num) == 9:
            receipt_num = "0" + str(receipt_num)

    #まとめ商品の確認ができたら、まとめ取り引き仕様のpdf作成とスプレッドシートに書き込み
    rapidnofollow = soup.find(class_ = "rapidnofollow")
    
    if rapidnofollow == None:
        save_name = str(mail_id)
        screenshot(driver, save_name)

        #エラーが出たら、単品商品の発送処理
        item_name = soup.find(class_ = 'decItmName').get_text(strip=True)
        item_name_list = format_item_name(item_name)

        #pdfの作成とスプレッドシートに書き込み
        single_item_pdf_spreadsheet(df, mail_id, item_name_list, delivery_method, delivery_time, receipt_num, item_password, src)

        eel.view_log_js('{}が作成されました\n'.format(save_name + ".pdf"))

    else:
        save_name = "まとめて取引_" + str(mail_id)
        screenshot(driver, save_name)

        #pdfの作成とスプレッドシートに書き込み
        multiple_items_pdf_spreadsheet(driver, wait, origin_df, df, mail_id, delivery_method, delivery_time, receipt_num, item_password, src)

        eel.view_log_js('{}が作成されました\n'.format(save_name + ".pdf"))

def main():

    #メインスプレッドシートのデータの読み込み
    origin_df = df_main_spreadsheet()

    #配送に必要な情報を整理
    df = origin_df[(origin_df['Purchaser'] != "pony13795") & (origin_df['Paid_Status'] != "") & (origin_df["Ship_Status"] == "") & (origin_df["Receive_Status"] == "") & (origin_df["Label_Process"] == "") & (origin_df["Ship_Order_No"] == "") & (origin_df["Evaluation_Proces"] == "")]
    # df = origin_df[(origin_df['Purchaser'] == "pony13795") & (origin_df['Paid_Status'] != "") & (origin_df["Ship_Status"] == "") & (origin_df["Receive_Status"] == "") & (origin_df["Label_Process"] == "") & (origin_df["Ship_Order_No"] == "") & (origin_df["Evaluation_Proces"] == "")]

    #新規発送がある時の処理
    if not df.empty:
        i = 0
        sagawa_count = 0

        #driverの設定
        path = os.getcwd()
        driver_path = ChromeDriverManager(path=path).install()

        for index, row in df.iterrows():

            try:
                #ドライバーの設定
                driver = set_driver(driver_path)
                driver.implicitly_wait(5)
                wait = WebDriverWait(driver, 5)

                # urlの取得
                url = row["Navi"]

                #try:
                i += 1

                #アクセスユーザー名とパスワードの取得
                user, password = format_url(url)

                #urlにアクセス
                driver.get(url)

                #ログイン
                login(driver, wait, user, password)

                #商品配送処理
                before_create_pdf_process(driver, wait, url)

                try:
                    #pdfの作成とスプレッドシートに書き込み
                    create_pdf_write_spreadsheet(driver, wait, origin_df, df, row)

                except:
                    index = int(index)
                    
                    try:
                        #まとめ取り引きの商品の確認
                        rapidnofollow = driver.find_element_by_class_name("rapidnofollow")
                        url = rapidnofollow.get_attribute("href")

                        driver.get(url)

                        #何をしているのか確認
                        while True:
                            try:
                                libJsExpandToggleBtn = driver.find_element_by_class_name('libJsExpandToggleBtn')
                                a = libJsExpandToggleBtn.find_element_by_tag_name('a')
                                a.click()
                            except NoSuchElementException:
                                break

                        html = driver.page_source.encode('utf-8')
                        soup = BeautifulSoup(html, 'lxml')

                        decBundle = soup.find(class_ = "decBundle")
                        item_list = decBundle.find_all('tr')

                        for j, item in enumerate(item_list):
                            #商品名
                            decItmName = item.find(class_ = "decItmName")

                            #商品の行を取得するための操作
                            href = decItmName.find("a")['href']
                            end_day = item.find(class_ = "decMDT").get_text(strip=True)
                            end_day = change_date_fetche_spreadsheet(end_day)
                            # print(end_day)

                            row = origin_df.query('Reference.str.contains("{}") & Date.str.contains("{}")'.format(href, end_day))
                            index = row.index[0]

                            if j == 0:
                                #スプレッドシートに書き込み
                                write_main_spreadsheet("AC{}".format(index + 2), "佐川急便")

                            else:
                                #スプレッドシートに書き込み
                                write_main_spreadsheet("Q{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("R{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("S{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("W{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("X{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("Y{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("Z{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("AA{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("AB{}".format(index + 2), "まとめて")
                                write_main_spreadsheet("AC{}".format(index + 2), "まとめて")

                    except:
                        write_main_spreadsheet("AC{}".format(index + 2), "佐川急便")

                    line_notification('ラベルツール\n' + '佐川急便の商品です\n' + url)
                    # line_error_notification('ラベルツール\n' + '佐川急便の商品です\n' + url)
                    sagawa_count += 1

                # driver.quit()

            except:
                i -= 1
                t = traceback.format_exc()
                line_error_notification('ラベルツール')
                t_list = [t[i: i + 1000] for i in range(0, len(t), 1000)]
                for message in t_list:
                    line_error_notification(message)

                eel.view_log_js('予期せぬエラーが発生しました。\n')
                eel.view_log_js('商品URL: {}'.format(url))
                
                line_error_notification('予期せぬエラーが発生しました。')
                line_error_notification('商品URL: {}'.format(url))

                continue

        
        i -= sagawa_count

        if i == 0:
            line_notification("新規の配送依頼はありませんでした")
            # line_error_notification("新規の配送依頼はありませんでした")
            eel.view_log_js('終了しました\n')

        else:
            line_notification("pdfが{}つ作成されました".format(str(i)))
            # line_error_notification("pdfが{}つ作成されました".format(str(i)))
            eel.view_log_js('終了しました\n')

    else:
        line_notification("新規の配送依頼はありませんでした")
        # line_error_notification("新規の配送依頼はありませんでした")
        eel.view_log_js('終了しました\n')
