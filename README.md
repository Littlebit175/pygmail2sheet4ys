# 概要
とあるECサイトの注文確認メールの内容をGmailから読み取り、スプレッドシートに転記するスクリプト

# 使い方
https://qiita.com/drafts/026f17aeb9261f7b04eb/

## 必要なPythonライブラリのインストール
```
$ pip3 install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

## config.jsonの編集
### start_date, end_date
検索対象メールの受信日の範囲を指定する。
Gmailの検索クエリの仕様(before, after)と同じ。

例）11/5に受信したメールを検索対象にしたい場合
```
    "start_date": "2024/11/5",
    "end_date": "2024/11/6",
```

→ 11/5のメールが検索にヒットする

```
    "start_date": "2024/11/5",
    "end_date": "2024/11/5",
``` 

→ 11/5のメールが検索にヒットしない

### spreadsheet_id
転記先スプレッドシートのID。
URLが`https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/`なら
```
    "spreadsheet_id": "YOUR_SPREADSHEET_ID",
```

### sheet_name
転記先スプレッドシートのシート名。
シート名が`Sheet1`なら
```
    "sheet_name": "Sheet1"
```

## スクリプト実行
```
$ python3 pygmail2sheet4ys.py
```