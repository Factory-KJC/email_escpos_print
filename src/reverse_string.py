from escpos.printer import Network
import socket

class TextReplacer():
    def __init__(self):
        # 置き換える文字のマッピングを定義
        self.replacements = {
            '\xa0': ' ',  #ノーブレークスペース
            '‘': "'",  # 左シングルクォート
            '’': "'",  # 右シングルクォート
            '“': '"',  # 左ダブルクォート
            '”': '"',  # 右ダブルクォート
            '—': '-',  # エムダッシュ
            '–': '-',  # エンダッシュ
        }

    def replace_in_list(self, text_list):
        """
        リスト内の文字列に対して、指定された置き換えを行うメソッド
        """
        # リスト内の各文字列に対して置き換えを行う
        return [
            ''.join(self.replacements.get(char, char) for char in text)
            for text in text_list
        ]

class ReverseNetworkPrinter(Network):
    def __init__(self, ip):
        # テキストバッファ（文字列のリスト）
        super().__init__(ip)
        self.text_buffer = []

    def add_text_to_buffer(self, text):
        # テキストを改行コードで分割
        text = self._limit_consecutive_newlines(text)
        lines = text.split('\n')

        for line in lines:
            self._add_single_line_to_buffer(line)

    def _limit_consecutive_newlines(self, text):
        """連続する改行を2回に制限するメソッド"""
        import re
        # CRLF (\r\n) を LF (\n) に統一
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        # 複数回の改行を制限（正規表現を使用して3回以上の連続改行を2回に置換）
        return re.sub(r'\n{3,}', '\n\n', text)

    def _add_single_line_to_buffer(self, text):
        # 12ピクセルの1バイト文字と24ピクセルの2バイト文字を含む文字列を
        # 1行が384ピクセルを超えないように分割し、逆順にしてバッファに追加する
        line = ""
        current_width = 0
        max_width = 384
        for char in text:
            if len(char.encode('shift_jis', errors='replace')) == 1:
                char_width = 12  # 1バイト文字（12ピクセル）
            else:
                char_width = 24  # 2バイト文字（24ピクセル）

            if current_width + char_width > max_width:
                # 384ピクセルを超えたら、新しい行として追加
                self.text_buffer.insert(0, line)
                line = char
                current_width = char_width
            else:
                line += char
                current_width += char_width

        # 最後の行を追加
        if line:
            self.text_buffer.insert(0, line)


    def encode_buffer_to_shift_jis(self):
        replacer = TextReplacer()
        converted_buffer = replacer.replace_in_list(self.text_buffer)
        # バッファに保存されたテキストをShift-JISでエンコード
        encoded_texts = [line.encode('shift_jis', errors='replace') for line in converted_buffer]
        return encoded_texts

    def print_encoded_text(self):
        try:
            # 左マージンを設定
            self._raw(b'\x1D\x4C\xC0\x00')  # 左マージンを24mm(192dot)に設定 GS L

            # 日本語用のコードページを設定（Shift JIS）
            self._raw(b'\x1B\x52\x08')  # 国際文字セットを「日本」へ
            self._raw(b'\x1B\x74\x01')  # 拡張ASCIIテーブルを「カタカナ」へ
            self._raw(b'\x1C\x43\x01')  # 文字コードを「Shift_JIS」へ

            # ESC { 1: 上下反転コマンド
            self._raw(b'\x1B\x7B\x01')

            for encoded_text in self.encode_buffer_to_shift_jis():
                self._raw(encoded_text)
                self._raw(b'\x1D\x72\x00')  # 印字完了待機コマンド
                self._raw(b'\n')  # 改行

        except socket.gaierror as e:
            print(f"プリンターのネットワークエラーが発生しました: {e}")
        except Exception as e:
            print(f"印刷中にエラーが発生しました: {e}")

        finally:
            try:
                self._raw(b'\x1B\x7B\x00') # 上下反転の解除 ESC { 0
                self._raw(b'\x1B\x40') # プリンターの初期化 ESC @
                self.close()
            except Exception as e:
                print(f"プリンターの接続を閉じる際にエラーが発生しました: {e}")
