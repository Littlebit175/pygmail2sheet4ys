# Copyright (c) 2024 Littlebit175

import os
import re
import base64
import datetime
import json
import pytz
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# 認証情報の設定
SCOPES = ['https://www.googleapis.com/auth/gmail.modify',  # 既読処理するためreadonlyではなくmodifyの権限が必要
          'https://www.googleapis.com/auth/spreadsheets']

def authenticate_gmail():
    creds = None
    # トークンファイルが存在するか確認
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # トークンが無効な場合、新しいトークンを取得
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Error refreshing token: {e}")
                creds = None
        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            # 新しいトークンを保存
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def authenticate_sheets():
    creds = None
    # トークンファイルが存在するか確認
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # トークンが無効な場合、新しいトークンを取得
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Error refreshing token: {e}")
                creds = None
        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            # 新しいトークンを保存
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
    return build('sheets', 'v4', credentials=creds)

def search_emails(service, start_date, end_date):
    jst = pytz.timezone('Asia/Tokyo')
    start_date_jst = jst.localize(datetime.datetime.strptime(start_date, '%Y/%m/%d'))
    end_date_jst = jst.localize(datetime.datetime.strptime(end_date, '%Y/%m/%d'))
    
    start_date_unix = int(start_date_jst.timestamp())
    end_date_unix = int(end_date_jst.timestamp())
    
    query = f'from:shopping-order-master@mail.yahoo.co.jp after:{start_date_unix} before:{end_date_unix}'
    results = service.users().messages().list(userId='me', q=query).execute()
    messages = results.get('messages', [])
    return messages

def extract_data_from_email(service, message_id):
    message = service.users().messages().get(userId='me', id=message_id).execute()
    payload = message['payload']
    headers = payload['headers']
    parts = payload.get('parts', [])
    data = {}
    subject = None
    recipient_email = None

    # メールタイトルを取得
    for header in reversed(headers):    # メールのヘッダを後方の要素から検索したほうが速い
        if header['name'] == 'Subject':
            subject = header['value']
            break

    # 宛先メールアドレスを取得
    for header in reversed(headers):    # メールのヘッダを後方の要素から検索したほうが速い
        if header['name'] == 'To':
            recipient_email = header['value']
            break

    # 注文確認メールの場合
    if subject.startswith('【Yahoo!ショッピング】ご注文の確認'):
        data['宛先メールアドレス'] = recipient_email
        data['メールタイトル'] = subject
        # メール本文を取得
        for part in parts:
            if part['mimeType'] == 'text/plain':
                body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                data.update(parse_email_body(body))
                break
    # 注文キャンセルメールの場合
    elif subject.startswith('【Yahoo!ショッピング】ご注文のキャンセル'):
        data['宛先メールアドレス'] = recipient_email
        data['メールタイトル'] = subject
        # メール本文を取得
        body = base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8')
        data.update(parse_email_body(body))
    # その他のメールの場合は空の辞書を返す
    else:
        return {}

    return data

def parse_email_body(body):
    data = {}
    patterns = {
        '注文ID': r'注文ID\s*：(.+)',
        'ストア名': r'ストア名：(.+)',
        '注文日': r'注文日時：(.+)',
        '商品の合計金額': r'商品の合計金額：\s*([\d,]+)円',
        'クーポン利用': r'クーポン利用：\s*-([\d,]+)円',
        '送料': r'送料：\s*([\d,]+)円',
        '商品券利用': r'商品券利用：\s*-([\d,]+)円',
        '支払方法': r'支払方法\s*：(.+)',
        'PayPay残高等利用額': r'PayPay残高等利用額：\s*-([\d,]+)円',
        'クレジットカード利用額': r'\n合計金額：\s*([\d,]+)円',
        '商品名': r'（1）(.+)\n'
    }

    # パターンにマッチする部分を抽出
    for key, pattern in patterns.items():
        match = re.search(pattern, body)
        if match:
            # 金額の文字列から「円」を削除し、カンマを削除して数値に変換
            data[key] = match.group(1).replace('円', '').replace(',', '')
        else:
            data[key] = ''

    # 注文日をYYYY/MM/DD形式に変換
    if '注文日' in data:
        data['注文日'] = datetime.datetime.strptime(data['注文日'], '%Y年%m月%d日 %H時%M分%S秒').strftime('%Y/%m/%d')

    return data

def append_to_sheet(service, spreadsheet_id, sheet_name, data_list):
    sheet = service.spreadsheets()
    
    # ヘッダ行を定義
    headers = [
        '注文ID', 'ストア名', '宛先メールアドレス', 'メールタイトル', '注文日',
        '商品の合計金額', 'クーポン利用', '送料', '商品券利用', '支払方法',
        'PayPay残高等利用額', 'クレジットカード利用額', '商品名'
    ]
    
    # ヘッダ行を追加
    header_body = {
        'values': [headers]
    }
    sheet.values().update(
        spreadsheetId=spreadsheet_id, 
        range=f'{sheet_name}!A1:M1', 
        valueInputOption='RAW', 
        body=header_body
    ).execute()
    
    # データ行を追加
    body = {
        'valueInputOption': 'RAW',
        'data': [
            {
                'range': f'{sheet_name}!A{index+2}:M{index+2}',
                'values': [
                    [
                        data.get('注文ID', ''), data.get('ストア名', ''), data.get('宛先メールアドレス', ''), 
                        data.get('メールタイトル', ''), data.get('注文日', ''), data.get('商品の合計金額', ''),
                        data.get('クーポン利用', ''), data.get('送料', ''), data.get('商品券利用', ''),
                        data.get('支払方法', ''), data.get('PayPay残高等利用額', ''), data.get('クレジットカード利用額', ''),
                        data.get('商品名', '')
                    ]
                ]
            } for index, data in enumerate(data_list)
        ]
    }
    sheet.values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def mark_messages_as_read(service, message_ids):
    """複数のメールを一括で既読にする"""
    try:
        # 最大で1000件ずつバッチ処理
        for i in range(0, len(message_ids), 1000):
            batch = message_ids[i:i + 1000]
            service.users().messages().batchModify(
                userId='me',
                body={
                    'ids': batch,
                    'removeLabelIds': ['UNREAD']
                }
            ).execute()
    except Exception as e:
        print(f"Error marking messages as read: {e}")

def main(config):
    gmail_service = authenticate_gmail()
    sheets_service = authenticate_sheets()
    messages = search_emails(gmail_service, config['start_date'], config['end_date'])

    data_list = []
    message_ids = []  # 既読にするメールのID一覧
    for message in messages:
        data = extract_data_from_email(gmail_service, message['id'])
        if data:  # データが空でないことを確認
            data_list.append(data)
            message_ids.append(message['id'])

    # 処理対象のメールを一括で既読にする
    if message_ids:
        mark_messages_as_read(gmail_service, message_ids)

    # "注文日" で昇順に並び替え、さらに "注文ID" で昇順に並び替え
    data_list = sorted(data_list, key=lambda x: (x.get('注文日', ''), x.get('注文ID', '')))

    append_to_sheet(sheets_service, config['spreadsheet_id'], config['sheet_name'], data_list)

if __name__ == '__main__':
    try:
        with open('config.json', 'r') as config_file:
            config = json.load(config_file)
        main(config)
    except FileNotFoundError:
        print("Error: 'config.json' file not found.")
    except json.JSONDecodeError:
        print("Error: 'config.json' contains invalid JSON.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")