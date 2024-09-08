# email_escpos_print
IMAPで新規受信したメールをESC/POS対応のサーマルプリンターで印刷するためのスクリプト

## 詳細
80mm用の機種を縦置きし，58mmの用紙を利用するために作成されています．
それ以外の条件で利用する場合は動作を保証できません．

## 設定
### プリンター設定
サーマルプリンターのipアドレスを入力してください．
```
# EPSONプリンターのネットワーク設定
p = Network("thermal.printer.ip.address")
```

### メールアドレス設定
メールサーバー，メールアドレス，パスワードを入力してください．
```
# メールのアカウント情報
username = "username@example.com"
password = "YourPassword"
mail_server = "imap.example.com"
```

### 定期実行
Windowsの場合はタスクスケジューラー，Linuxの場合はcronを利用して定期実行するとメールを受信するたびに印刷できます．
