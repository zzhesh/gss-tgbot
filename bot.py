import asyncio
import pickle
import telebot
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1NDJiEtxnISwi7DumpYX9vmuFK2RKsnCYrCqcfwVd-JU'
TOKEN = '1497294051:AAH8n6X4A1Cfa8wu6aYHr90G5ZXhokyIxos'
bot = telebot.TeleBot(TOKEN)
url = 'https://docs.google.com/spreadsheets/d/1NDJiEtxnISwi7DumpYX9vmuFK2RKsnCYrCqcfwVd-JU'



@bot.message_handler(content_types=['text'])
def tg(message):
       
    if message.text == ('/start'):               # приветствие и примеры
        bot.send_message(message.from_user.id, f'url: {url}\nсинтаксис:\n/gss адресс\nзначение ячейки|значение ячейки\nyиже отправлю пару примеров') 
        bot.send_message(message.from_user.id, '/gss C1\nThis is C1')
        bot.send_message(message.from_user.id, '/gss A2:C3\na2|b2|c2\na3|b3|c3')
        bot.send_message(message.from_user.id, '/gss A1:B7\n1|100\n2|465\n3|54\n4|87\n5|134\n6|879\nИтого|=B1+B2+B4+B5+B6')
            
    elif message.text.startswith('/gss'):        # команда для едита
        msg = str(message.text)[5:]
         
            
        if 'rows' in msg: majorDimension = 'ROWS'            #
        elif 'columns' in msg: majorDimension = 'COLUMNS'    # орентация заполнения таблички
        else: majorDimension = 'ROWS'                        #
    
        range = msg.split('\n')[0].split(' ')[0] #  адреса ячеек
        values = []                              #
        for i in msg.split('\n')[1:]:            #  значения ячеек
            values.append(i.split('|'))          #
    
        bot.send_message(message.from_user.id,   # фидбек в чат
                         f'range: {range}\nvalues: {values}')

        try:   
            r = sheet.values().batchUpdate( 
                spreadsheetId=SPREADSHEET_ID,
                body = {'value_input_option' : 'USER_ENTERED',
                        'data': {"range": range,
                                 'majorDimension': majorDimension,
                                 'values': values}}).execute()            
        except Exception as e:
            (f'in batchUpdate >>\t{e}\n')

# логин в гуглах
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

print('in')
bot.polling(none_stop=True, interval=0)
