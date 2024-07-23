import extract_msg
import os
import shutil
import re
import time
from email.message import EmailMessage
from email.policy import default
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

# MSGファイルが格納されているフォルダを入力
msg_folder = input("MSGファイルが格納されているフォルダを入力してください: ")

# フォルダの存在確認
if not os.path.isdir(msg_folder):
    print("指定されたフォルダが存在しません。プログラムを終了します。")
    input("Enterキーを押して終了してください...")
    exit()

tmp_folder = os.path.join(msg_folder, 'tmp')
eml_folder = os.path.join(msg_folder, 'emlFiles')

# フォルダの作成
os.makedirs(tmp_folder, exist_ok=True)
os.makedirs(eml_folder, exist_ok=True)

# タイトルに日付を追加するか確認
add_date_to_title = input("タイトルにメール受信日時を追加しますか？（y/n）: ").strip().lower()

# 変換済みファイルのリストを保存するファイルパス
converted_files_list_path = os.path.join(eml_folder, 'converted_files.txt')

# 既に変換済みのファイルをリストから読み込む
if os.path.exists(converted_files_list_path):
    with open(converted_files_list_path, 'r') as file:
        converted_files = set(file.read().splitlines())
else:
    converted_files = set()

# MSGフォルダ内のすべてのMSGファイルを取得
msg_files = [f for f in os.listdir(msg_folder) if f.endswith('.msg')]
total_files = len(msg_files)

# 変換されたファイル数をカウント
converted_count = 0
# スキップされたファイル数をカウント
skipped_count = 0

# エラーログファイルの設定
error_log_path = os.path.join(msg_folder, 'error.log')

for msg_file in msg_files:
    if msg_file in converted_files:
        skipped_count += 1
        print(f"スキップ: {msg_file}（既に変換済み）")
        continue

    try:
        # MSGファイルの読み込み
        msg_path = os.path.join(msg_folder, msg_file)
        msg = extract_msg.Message(msg_path)

        # メッセージ部分の情報取得
        subject = sanitize_filename(msg.subject)
        sender = msg.sender
        date = msg.date.strftime('%Y%m%d%H%M%S')  # 日付を文字列に変換

        # 受信者情報を "名前<メールアドレス>" 形式に変換
        recipients = [f"{recipient.name} <{recipient.email}>" for recipient in msg.recipients]
        body = msg.body

        # タイトルに日付を追加
        if add_date_to_title == 'y':
            eml_title = f"{date}_{subject}.eml"
        else:
            eml_title = f"{subject}.eml"

        # EMLメッセージの作成（マルチパート）
        eml_msg = MIMEMultipart()
        eml_msg['Subject'] = subject
        eml_msg['From'] = sender
        eml_msg['To'] = ', '.join(recipients)
        eml_msg['Date'] = msg.date.strftime('%a, %d %b %Y %H:%M:%S %z')

        # メッセージ本文の追加
        eml_msg.attach(MIMEText(body, 'plain'))

        # 添付ファイルの処理と格納
        for attachment in msg.attachments:
            attachment.save(customPath=tmp_folder)
            file_path = os.path.join(tmp_folder, attachment.longFilename)

            # 添付ファイルが存在することを確認
            while not os.path.exists(file_path):
                time.sleep(0.1)

            # 添付ファイルをEMLに追加
            with open(file_path, 'rb') as f:
                mime_part = MIMEBase('application', 'octet-stream')
                mime_part.set_payload(f.read())
                encoders.encode_base64(mime_part)
                mime_part.add_header(
                    'Content-Disposition',
                    f'attachment; filename={attachment.longFilename}',
                )
                eml_msg.attach(mime_part)

        # EMLファイルの保存
        eml_file_path = os.path.join(eml_folder, eml_title)
        with open(eml_file_path, 'wb') as eml_file:
            eml_file.write(eml_msg.as_bytes(policy=default))

        # 変換されたファイルをリストに追加
        converted_files.add(msg_file)

        # 一時フォルダ内の添付ファイルを削除
        for file in os.listdir(tmp_folder):
            try:
                os.remove(os.path.join(tmp_folder, file))
            except PermissionError as e:
                print(f"ファイル {file} の削除中にエラーが発生しました: {e}")
                with open(error_log_path, 'a', encoding='utf-8') as error_log:
                    error_log.write(f"ファイル {msg_file} の変換中にエラーが発生しました: {e}\n")

        # 削除が完了したことを確認
        while os.listdir(tmp_folder):
            time.sleep(0.1)

        # 変換されたファイルのカウントを増加
        converted_count += 1

        # 進捗を表示
        print(f"変換中: {converted_count}/{total_files} - {msg_file}")

    except Exception as e:
        with open(error_log_path, 'a', encoding='utf-8') as error_log:
            error_log.write(f"ファイル {msg_file} の変換中にエラーが発生しました: {e}\n")
            # 詳細なエラーログ
            import traceback
            error_log.write(f"詳細エラートレースバック: {traceback.format_exc()}\n")

# 変換済みファイルのリストを保存
with open(converted_files_list_path, 'w', encoding='utf-8') as file:
    for converted_file in converted_files:
        file.write(f"{converted_file}\n")

print(f"全てのMSGファイルがEMLに変換されました。変換されたファイル数: {converted_count}/{total_files - skipped_count}")
print(f"スキップされたファイル数: {skipped_count}")

# エラーログの内容をユーザーに通知
if os.path.exists(error_log_path):
    print(f"エラーログファイルが {error_log_path} に作成されました。詳細はログファイルを確認してください。")

# 終了前にユーザー入力を待つ
input("Enterキーを押して終了してください...")
