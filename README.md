# email_escpos_print
IMAPで新規受信したメールをESC/POS対応のサーマルプリンターで印刷するためのスクリプト

## 詳細
80mm用の機種を縦置きし，58mmの用紙を利用するために作成されています．
それ以外の条件で利用する場合は動作を保証できません．
EPSON TM-T90Ⅱのみで動作確認しています．

## 設定
### プリンター設定
サーマルプリンターのipアドレスを入力してください．
```
# EPSONプリンターのネットワーク設定
PRINTER_IP_ADDRESS = "thermal.printer.ip.address"
```

### メールアドレス設定
`account.cfg`にメールサーバー，メールアドレス，パスワードを入力してください．

パスワード等に一部の記号が使われている場合(%，$等)，%%，$$等で対処してください．
Gmail，iCloud等で，二要素認証を利用している場合には，アプリパスワードを生成して利用してください．

複数のアカウントを利用する場合は以下のように適宜追加してください．
```
[Account1]
email = username@example.com
password = YourPassword
mail_server = imap.example.com

[Account2]
email = username2@example.com
password = YourPassword2
mail_server = imap.example.com

[Account2]
email = username3@example.org
password = YourPassword3
mail_server = imap.example.org

....
```

### 定期実行
Windowsの場合はタスクスケジューラー，Linuxの場合はcronを利用して定期実行するとメールを受信するたびに印刷できます．
