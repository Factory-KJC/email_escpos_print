import imaplib
import email
import os
from escpos.printer import Network
from escpos.printer import Escpos

# 最後に処理したメールIDを保存するファイル
LAST_ID_FILE = 'last_email_id.txt'

def get_last_email_id():
    if os.path.exists(LAST_ID_FILE):
        with open(LAST_ID_FILE, 'r') as file:
            return file.read().strip()
    return None

def set_last_email_id(email_id):
    with open(LAST_ID_FILE, 'w') as file:
        file.write(email_id)

def fetch_and_print_email():
    # メールサーバーに接続
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login('your_email@gmail.com', 'your_password')

    # 受信トレイを選択
    mail.select('inbox')

    # 最後に処理したメールIDを取得
    last_email_id = get_last_email_id()

    # 全てのメールを検索
    result, data = mail.search(None, 'ALL')

    # メールIDリストを取得
    email_ids = data[0].split()

    # 処理すべきメールIDを選択
    new_email_ids = [email_id for email_id in email_ids if last_email_id is None or email_id > last_email_id]

    if not new_email_ids:
        print("新しいメールはありません。")
        return

    for email_id in new_email_ids:
        result, data = mail.fetch(email_id, '(RFC822)')
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        # タイトルと送信者を抽出
        subject = msg['subject']
        sender = msg['from']

        # プリンターのIPアドレスを指定
        network_printer = Network("192.168.1.100")

        # プリンターに接続
        p = Escpos(network_printer)

        # メール情報を印刷
        p.text("Subject: {}\n".format(subject))
        p.text("From: {}\n".format(sender))

        # 紙をカット
        p.cut()

        # 最後に処理したメールIDを更新
        set_last_email_id(email_id)

    # セッションを終了
    mail.logout()

# メイン関数の実行
if __name__ == "__main__":
    fetch_and_print_email()
