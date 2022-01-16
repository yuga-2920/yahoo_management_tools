import gspread
from oauth2client.service_account import ServiceAccountCredentials 
import pandas as pd


#メインスプレッドシートの設定
def set_main_spreadsheet(): #=>worksheet

    #リフレッシュトークンの発行
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    #認証情報の設定
    credentials = ServiceAccountCredentials.from_json_keyfile_name('yahuoku-shipping-evaluation-dc1a8cb8daf4.json', scope)
    gc = gspread.authorize(credentials)

    #スプレッドシートの読み込み
    SPREADSHEET_KEY = "1sFNfVTyCFO52esE1Pxej_aMffKjWNHZtTB3hOZz0xZ8"
    workbook = gc.open_by_key(SPREADSHEET_KEY)
    worksheet = workbook.worksheet('alarm')

    return worksheet


#メインスプレッドシートから読み込みdataframeに整形
def df_main_spreadsheet():

    #worksheetの取得
    worksheet = set_main_spreadsheet()

    #dataframe化
    cell_list = worksheet.get_all_values()
    df = pd.DataFrame(cell_list[1:], columns=cell_list[0])

    return df


#メインワークシートの書き込み
def write_main_spreadsheet(row_column, date):

    #worksheetの取得
    worksheet = set_main_spreadsheet()

    #書き込み
    worksheet.update_acell(row_column, date)
