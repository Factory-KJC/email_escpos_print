import imaplib
import email
from email.header import decode_header
from PIL import Image, ImageDraw, ImageFont
from escpos.printer import Network
import os
from html import unescape
from bs4 import BeautifulSoup
import chardet
from urllib.parse import unquote

# 処理済みメールIDを保存するファイル
PROCESSED_MAILS_FILE = "processed_mails.txt"
MAX_TEXT_LENGTH = 300  # メール本文の最大文字数

# ヘッダー情報をデコードしてUTF-8に変換する関数
def decode_mime_words(s):
    decoded_fragments = decode_header(s)
    return ''.join(
        str(fragment if isinstance(fragment, str) else fragment.decode(encoding or 'utf-8'))
        for fragment, encoding in decoded_fragments
    )

# 処理済みメールIDを取得
def get_processed_mail_ids():
    if os.path.exists(PROCESSED_MAILS_FILE):
        with open(PROCESSED_MAILS_FILE, "r") as f:
            return set(line.strip() for line in f)  # 空白を削除してIDをセットにする
    return set()

# 処理済みメールIDを保存
def save_processed_mail_id(mail_id):
    with open(PROCESSED_MAILS_FILE, "a") as f:
        f.write(f"{mail_id.strip()}\n")  # 空白を削除してIDを保存

def extract_text_from_html(html_content):
    """ HTMLメールからテキストを抽出する関数 """
    try:
        if isinstance(html_content, bytes):
            # バイトデータから文字列に変換
            detected_encoding = chardet.detect(html_content)['encoding']
            if detected_encoding is None:
                detected_encoding = 'utf-8'
            html_content = html_content.decode(detected_encoding, errors='ignore')

        # URLエンコードをデコード
        html_content = unquote(html_content)

        soup = BeautifulSoup(html_content, 'html.parser')
        text = soup.get_text(separator='\n', strip=True)
        return unescape(text)
    except Exception as e:
        print(f"HTMLからテキストを抽出する際にエラーが発生しました: {e}")
        return "HTMLメールの内容を抽出できませんでした"

# メールを取得する関数
def fetch_emails(username, password, mail_server, folder="INBOX"):
    mail = imaplib.IMAP4_SSL(mail_server)
    mail.login(username, password)
    mail.select(folder)

    processed_mail_ids = get_processed_mail_ids()
    status, messages = mail.search(None, 'UNSEEN')
    mail_ids = messages[0].split()

    if not mail_ids:
        print("未読のメールはありません")
        return []

    emails_to_process = []

    for mail_id in mail_ids:
        mail_id_str = mail_id.decode('utf-8')
        if mail_id_str not in processed_mail_ids:
            status, msg_data = mail.fetch(mail_id, "(BODY.PEEK[])")

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    subject = decode_mime_words(msg["Subject"])
                    from_ = decode_mime_words(msg.get("From"))

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

                    # 処理済みとしてIDを保存
                    save_processed_mail_id(mail_id_str)
                    emails_to_process.append((subject, from_, body))

    return emails_to_process

#画像を180度回転する関数
def rotete_image_180(image):
    return image.rotate(180, expand=True)

def wrap_text(text, font, max_width):
    """
    テキストを指定した幅に収まるように自動改行する関数。
    """
    lines = []
    current_line = ""

    for char in text:
        # 現在の行に文字を追加して描画幅を計測
        test_line = current_line + char
        width, _ = font.getbbox(test_line)[2:4]  # 幅を取得

        if width <= max_width:
            # 幅内に収まる場合は、文字を追加
            current_line = test_line
        else:
            # 幅を超えた場合は行を確定して新しい行を開始
            lines.append(current_line)
            current_line = char  # 現在の文字を次の行に追加

    # 最後の行を追加
    if current_line:
        lines.append(current_line)

    return '\n'.join(lines)


# テキストを画像に変換する関数
def text_to_image(text, width=384):
    # 日本語フォントのロード

    font_path = "fonts/NotoSansCJK-Regular.ttc"
    font_size = 20 # フォントサイズ
    font = ImageFont.truetype(font_path, font_size)

    # CRLF（\r\n）をLF（\n）に置き換える
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # テキストを改行ごとに処理して収まるように改行
    wrapped_text = ""
    for paragraph in text.split('\n'):
        wrapped_text += wrap_text(paragraph, font, width) + '\n'

    # テキストを行ごとに分割
    lines = wrapped_text.split('\n')

    # 1行の高さを計算
    _, line_height = font.getbbox("A")[2:4]  # テキスト高さを取得

    # 画像サイズを計算
    image_height = line_height * len(lines) + 5 * (len(lines) - 1)  # 行間5ピクセルを追加

    # 画像を作成
    image = Image.new('1', (width, image_height), 255)
    draw = ImageDraw.Draw(image)

    # テキストを描画
    y = 0
    for line in lines:
        draw.text((0, y), line, font=font, fill=0)
        y += line_height + 5  #行間5ピクセルを追加

    return image

# プリンターに画像を送信する関数
def print_image(image):

    # EPSONプリンターのネットワーク設定
    p = Network("thermal.printer.ip.address")

    # 印刷可能領域を設定（48mm幅、384ドット）
    # p._raw(b'\x1D\x57\x70\x01')  # 印刷幅を48mmに設定

    # 左マージンを設定
    p._raw(b'\x1D\x4C\x00\x00')  # 左マージンを0mm(0dot)に設定

    # 画像を上下反転
    rotated_image = rotete_image_180(image)

    # 画像を左端に寄せるためのキャンバスを作成
    canvas_width = 640  # 印刷幅（ドット）
    canvas_height = rotated_image.height
    canvas = Image.new('1', (canvas_width, canvas_height), 255)  # 白い背景
    # 48mm幅の画像を左端に配置
    image_x_offset = 0
    canvas.paste(rotated_image, (image_x_offset, 0))  # 左上に配置

    p.image(canvas)
    p.cut()

    p.close()

# メイン処理
if __name__ == "__main__":
    # メールのアカウント情報
    username = "username@example.com"
    password = "YourPassword"
    mail_server = "imap.example.com"

    # メールを取得
    emails_to_print = fetch_emails(username, password, mail_server)

    if emails_to_print:
        for subject, from_, body in emails_to_print:
            # メールをフォーマット
            email_text = f"Subject: {subject}\nFrom: {from_}\n\n{body}"

            # テキストを画像に変換
            image = text_to_image(email_text)

            # プリンターで印刷
            print_image(image)
    else:
        print("未読のメールがありません、またはすべての未読メールは既に印刷されています")
