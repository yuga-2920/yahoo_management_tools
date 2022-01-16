import requests
import datetime
import pytz

#LINEの設定
def line_error_notification(message):
    #TOKEN = '00RhRrvc8wlHphEVIs8P5ppPsfVGC08EfObKkwZ2A32'
    TOKEN = 'xeQAni6AH8QiR8A1s5BrS2ovdcUZ6nyR69GXkAbrVwR'
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
