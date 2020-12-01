import telebot
import asyncio
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request



SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1NDJiEtxnISwi7DumpYX9vmuFK2RKsnCYrCqcfwVd-JU'
TOKEN = '1497294051:AAH8n6X4A1Cfa8wu6aYHr90G5ZXhokyIxos'
bot = telebot.TeleBot(TOKEN)
url = 'https://docs.google.com/spreadsheets/d/1NDJiEtxnISwi7DumpYX9vmuFK2RKsnCYrCqcfwVd-JU'

creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()


try:

    @bot.message_handler(content_types=['text'])
    def tg(message):
        if message.text == ('/start'):
            bot.send_message(message.from_user.id,
            f'''url: {url}\nсинтаксис:
            /gss адресс
            значение ячейки|значение ячейки
ниже отправлю пару примеров''') # приветствие и примеры
            bot.send_message(message.from_user.id,
            f'/gss C1\nThis is C1')
            bot.send_message(message.from_user.id,
            f'''/gss A2:C3
a2|b2|c2
a3|b3|c3''')
            bot.send_message(message.from_user.id,
            f'''/gss A1:B7
1|100
2|465
3|54
4|87
5|134
6|879
Итого|=B1+B2+B4+B5+B6''')

            # получение значений из бота и отправка их на едит
        elif message.text.startswith('/gss'):
            msg = str(message.text)[5:]

            if 'rows' in msg: majorDimension = 'ROWS'
            elif 'columns' in msg: majorDimension = 'COLUMNS'
            else: majorDimension = 'ROWS'

            range = msg.split('\n')[0].split(' ')[0]

            values = []
            for i in msg.split('\n')[1:]:
                values.append(i.split('|'))

            bot.send_message(message.from_user.id,
                            f'range: {range}\nvalues: {values}')

            try:
                r = sheet.values().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body = {
                        'value_input_option' : 'USER_ENTERED',
                        'data': {"range": range,
                                'majorDimension': majorDimension,
                                'values': values,
                                }
                            }).execute()
            except Exception as e:
                print(e)
    print('in')
    bot.polling(none_stop=True, interval=0)

except Exception as e:
    print(e)
