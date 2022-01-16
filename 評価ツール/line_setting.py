import requests
import datetime
import pytz

#LINEの設定
def line_error_notification(message):
    TOKEN = '###'
    api_url = 'https://notify-api.line.me/api/notify'

    time = datetime.datetime.now(pytz.timezone('Asia/Tokyo'))
    time = time.strftime('%Y/%m/%d %H:%M:%S')

    TOKEN_dic = {
    'Authorization': 'Bearer' + ' ' + TOKEN
    }

    send_dic = {
        'message': time +  '\n' +message
    }

    requests.post(api_url, headers=TOKEN_dic, data=send_dic)
