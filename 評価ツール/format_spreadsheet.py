import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials 



#メインスプレッドシートの設定
def set_main_spreadsheet(): #=>worksheet

    #リフレッシュトークンの発行
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    #認証情報の設定
    credentials = ServiceAccountCredentials.from_json_keyfile_name('###.json', scope)
    gc = gspread.authorize(credentials)

    #スプレッドシートの読み込み
    SPREADSHEET_KEY = "###"
    workbook = gc.open_by_key(SPREADSHEET_KEY)
    worksheet = workbook.worksheet('alarm')

    return worksheet


#メインスプレッドシートから読み込みdataframeに整形
def df_main_spreadsheet(worksheet): # => dataframe, mail_id_list

    #dataframe化
    cell_list = worksheet.get_all_values()
    df = pd.DataFrame(cell_list[1:], columns=cell_list[0])

    return df


#メインワークシートの書き込み
def write_main_spreadsheet(worksheet, row_column, date):

    #書き込み
    worksheet.update_acell(row_column, date)

#値の取得
def get_main_spreadsheet(worksheet, row_column):

    #読み込み
    value = worksheet.acell(row_column).value

    return value
