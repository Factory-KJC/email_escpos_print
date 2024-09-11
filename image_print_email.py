import imaplib
import email
from email.header import decode_header
import socket
from PIL import Image, ImageDraw, ImageFont
from escpos.printer import Network
import os
from html import unescape
from bs4 import BeautifulSoup
import chardet
from urllib.parse import unquote
import configparser
from email.utils import parsedate_to_datetime
import time
import pytz

# 処理済みメールIDを保存するファイル
PROCESSED_MAILS_FILE = "processed_mails.txt"
MAX_TEXT_LENGTH = 300

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


def rotate_image_180(image):
    return image.rotate(180, expand=True)

def wrap_text(text, font, max_width):
    """ テキストを指定した幅に収まるように自動改行する関数 """
    lines = []
    current_line = ""

    for char in text:
        test_line = current_line + char
        width, _ = font.getbbox(test_line)[2:4]  # 幅を取得

        if width <= max_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = char

    if current_line:
        lines.append(current_line)

    return '\n'.join(lines)

def text_to_image(text, width=384):
    font_path = "fonts/NotoSansCJK-Regular.ttc"
    font_size = 20
    font = ImageFont.truetype(font_path, font_size)

    text = text.replace('\r\n', '\n').replace('\r', '\n')
    wrapped_text = ""
    for paragraph in text.split('\n'):
        wrapped_text += wrap_text(paragraph, font, width) + '\n'

    lines = wrapped_text.split('\n')
    _, line_height = font.getbbox("A")[2:4]
    image_height = line_height * len(lines) + 5 * (len(lines) - 1)
    image = Image.new('1', (width, image_height), 255)
    draw = ImageDraw.Draw(image)

    y = 0
    for line in lines:
        draw.text((0, y), line, font=font, fill=0)
        y += line_height + 5

    return image

def print_image(image):
    try:
        # EPSONプリンターのネットワーク設定
        p = Network("thermal.printer.ip.address")

        # 左マージンを設定
        p._raw(b'\x1D\x4C\x00\x00')  # 左マージンを0mm(0dot)に設定

        # 画像を上下反転
        rotated_image = rotate_image_180(image)

        # 画像を左端に寄せるためのキャンバスを作成
        canvas_width = 640  # 印刷幅（ドット）
        canvas_height = rotated_image.height
        canvas = Image.new('1', (canvas_width, canvas_height), 255)  # 白い背景
        # 48mm幅の画像を左端に配置
        image_x_offset = 0
        canvas.paste(rotated_image, (image_x_offset, 0))  # 左上に配置

        # 画像を送信
        p.image(canvas)
        # 印字完了を待つコマンド
        p._raw(b'\x1D\x72\x00')  # 印字完了待機コマンド
        # 用紙をカット
        p.cut()

        # 印字完了を待機
        time.sleep(5)

    except socket.gaierror as e:
        print(f"プリンターのネットワークエラーが発生しました: {e}")
    except Exception as e:
        print(f"印刷中にエラーが発生しました: {e}")
    finally:
        try:
            p.close()
        except Exception as e:
            print(f"プリンターの接続を閉じる際にエラーが発生しました: {e}")


if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("account.cfg")

    for account in config.sections():
        username = config[account]["email"]
        password = config[account]["password"]
        mail_server = config[account]["mail_server"]

        check_host_reachable(mail_server)  # ホストの解決をチェック

        emails_to_print = fetch_emails(account, username, password, mail_server)

        if emails_to_print:
            for subject, from_, body, received_time_str, email_address in emails_to_print:
                email_text = f"To: {email_address}\nReceived at: {received_time_str}\nSubject: {subject}\nFrom: {from_}\n\n{body}"
                image = text_to_image(email_text)
                print_image(image)
        else:
            print(f"{account}: 未読のメールがありません、またはすべての未読メールは既に印刷済みです。")
