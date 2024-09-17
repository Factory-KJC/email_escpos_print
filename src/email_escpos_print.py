from .reverse_string import ReverseNetworkPrinter
import imaplib
import email
from email.header import decode_header
import socket
import os
from html import unescape
from bs4 import BeautifulSoup
import chardet
from urllib.parse import unquote
import configparser
from email.utils import parsedate_to_datetime
import pytz

# 処理済みメールIDを保存するファイル
PROCESSED_MAILS_FILE = "../processed_mails.txt"
MAX_TEXT_LENGTH = 300
PRINTER_IP_ADDRESS = "thermal.printer.ip.address"

def main():
    config = configparser.ConfigParser()
    config.read("../account.cfg")

    p = ReverseNetworkPrinter(PRINTER_IP_ADDRESS)

    for account in config.sections():
        username = config[account]["email"]
        password = config[account]["password"]
        mail_server = config[account]["mail_server"]

        check_host_reachable(mail_server)  # ホストの解決をチェック

        emails_to_print = fetch_emails(account, username, password, mail_server)

        if emails_to_print:
            for email_to_print in emails_to_print:
                # ネストされたリスト/タプルの構造を確認
                subject, from_, body, received_time_str, email_address = email_to_print
                email_text = f"To: {email_address}\nReceived at: {received_time_str}\nSubject: {subject}\nFrom: {from_}\n\n{body}"

                p.add_text_to_buffer(email_text)

                print(email_text)

                # すべてのメールのバッファを追加し終えた後に印刷
                p.print_encoded_text()

                p.text_buffer.clear()

                p.cut()  # ここで一度だけカットを実行する
        else:
            print(f"{account}: 未読のメールがありません、またはすべての未読メールは既に印刷済みです。")


def decode_mime_words(s):
    decoded_fragments = decode_header(s)
    return ''.join(
        str(fragment if isinstance(fragment, str) else fragment.decode(encoding or 'utf-8'))
        for fragment, encoding in decoded_fragments
    )

def get_processed_mail_uids():
    """
    処理済みのメールUIDをファイルから読み込み、アカウント名とUIDのペアをセットにして返す関数。
    """
    processed_mail_uids = {}
    if os.path.exists(PROCESSED_MAILS_FILE):
        with open(PROCESSED_MAILS_FILE, "r") as f:
            for line in f:
                line = line.strip()
                if line:
                    account, uid = line.split(':', 1)  # UIDとアカウント名を分割
                    if account not in processed_mail_uids:
                        processed_mail_uids[account] = set()
                    processed_mail_uids[account].add(uid)
    return processed_mail_uids


def save_processed_mail_uid(account, mail_uid):
    """
    処理済みのメールUIDをファイルに保存する関数。
    すでに存在するUIDは保存しないようにする。
    """
    processed_mail_uids = get_processed_mail_uids()

    uid_with_account = f"{account}:{mail_uid.strip()}"
    if uid_with_account.split(':', 1)[1] not in processed_mail_uids.get(account, set()):
        with open(PROCESSED_MAILS_FILE, "a") as f:
            f.write(f"{uid_with_account}\n")  # UIDとアカウントを保存


def extract_text_from_html(html_content):
    """ HTMLメールからテキストを抽出する関数 """
    try:
        if isinstance(html_content, bytes):
            detected_encoding = chardet.detect(html_content)['encoding']
            if detected_encoding is None:
                detected_encoding = 'utf-8'
            html_content = html_content.decode(detected_encoding, errors='ignore')

        html_content = unquote(html_content)
        soup = BeautifulSoup(html_content, 'html.parser')
        text = soup.get_text(separator='\n', strip=True)
        return unescape(text)
    except Exception as e:
        print(f"HTMLからテキストを抽出する際にエラーが発生しました: {e}")
        return "HTMLメールの内容を抽出できませんでした"

def check_host_reachable(hostname):
    try:
        socket.gethostbyname(hostname)
        print(f"{hostname} は正常に解決されました。")
    except socket.error as e:
        print(f"{hostname} の解決に失敗しました: {e}")

def fetch_emails(account, username, password, mail_server, folder="INBOX"):
    try:
        print(f"接続中: {mail_server}")
        mail = imaplib.IMAP4_SSL(mail_server)
        mail.login(username, password)
        mail.select(folder)
        print(f"ログイン成功: {mail_server}")

    except imaplib.IMAP4.error as e:
        print(f"IMAP4エラーが発生しました: {e}")
        if "Invalid credentials" in str(e):
            print("認証情報が無効です。ユーザー名またはパスワードを確認してください。")
        return []
    except socket.gaierror as e:
        print(f"ネットワークエラーが発生しました: {e}")
        return []
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}")
        return []

    # 処理済みメールUIDを取得
    processed_mail_uids = get_processed_mail_uids()

    # 未読かつ削除されていないメールを取得 (UIDを含む)
    status, messages = mail.uid('search', None, '(UNSEEN NOT DELETED)')
    # ステータスを確認し、メッセージが存在するかチェック
    if status == 'OK' and messages[0]:
        mail_uids = messages[0].split()
    else:
        print("未読メールがありません、または検索に失敗しました")
        return []

    emails_to_process = []

    # 日本標準時(JST)タイムゾーンを取得
    jst = pytz.timezone('Asia/Tokyo')

    for mail_uid in mail_uids:
        mail_uid_str = mail_uid.decode('utf-8')

        # 処理済みUIDにないか確認
        if mail_uid_str not in processed_mail_uids.get(account, set()):
            try:
                status, msg_data = mail.uid('fetch', mail_uid, "(BODY.PEEK[])")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])

                        subject = decode_mime_words(msg["Subject"])
                        from_ = decode_mime_words(msg.get("From"))
                        date_header = msg["Date"]

                        # 受信日時をパース
                        try:
                            received_datetime = parsedate_to_datetime(date_header)
                            received_datetime_jst = received_datetime.astimezone(jst)
                            received_time_str = received_datetime_jst.strftime("%Y-%m-%d %H:%M:%S")
                        except Exception as e:
                            print(f"受信時刻のパースに失敗しました: {e}")
                            received_time_str = "不明"

                        body = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                content_disposition = str(part.get("Content-Disposition"))

                                # 添付ファイルをスキップ
                                if 'attachment' in content_disposition:
                                    continue

                                content_type = part.get_content_type()
                                charset = part.get_content_charset()

                                if charset is None:
                                    charset = 'utf-8'

                                if content_type == "text/plain":
                                    body += part.get_payload(decode=True).decode(charset, errors='ignore')
                                elif content_type == "text/html":
                                    html_body = part.get_payload(decode=True)
                                    body += extract_text_from_html(html_body)

                            if len(body) > MAX_TEXT_LENGTH:
                                body = body[:MAX_TEXT_LENGTH] + '... [内容が長すぎます]'
                        else:
                            # シングルパートのメールの場合
                            content_type = msg.get_content_type()
                            charset = msg.get_content_charset()
                            if charset is None:
                                charset = 'utf-8'
                            if content_type == "text/plain":
                                body = msg.get_payload(decode=True).decode(charset, errors='ignore')
                            elif content_type == "text/html":
                                html_body = msg.get_payload(decode=True)
                                body = extract_text_from_html(html_body)

                            if len(body) > MAX_TEXT_LENGTH:
                                body = body[:MAX_TEXT_LENGTH] + '... [内容が長すぎます]'

                        # 処理済みとしてUIDを保存
                        save_processed_mail_uid(account, mail_uid_str)
                        emails_to_process.append((subject, from_, body, received_time_str, username))

            except imaplib.IMAP4.error as e:
                print(f"メール取得中にIMAP4エラーが発生しました: {e}")
            except Exception as e:
                print(f"メール処理中に予期しないエラーが発生しました: {e}")

    return emails_to_process


if __name__ == "__main__":
    main()
