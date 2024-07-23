# Convert MSG to EML / MSGをEMLに変換

## :us: English Version

### Description
This Python program converts Microsoft Outlook MSG files to EML format while preserving email attachments and metadata.
It is designed to process multiple MSG files from a specified directory, convert them to EML format, and store the converted files in a separate directory. Additionally, 
the program offers an option to prefix the email subject with the received date and time, ensuring organized and chronological email records.

### Features
- Batch Conversion: Converts multiple MSG files in a directory to EML format.
- Attachment Handling: Extracts and includes all attachments from the MSG files in the converted EML files.
- Metadata Preservation: Retains essential email metadata such as sender, recipients, subject, and date.
- Date Prefix Option: Allows users to prefix the email subject with the received date and time for better organization.
- Temporary File Management: Manages temporary files during conversion and cleans up after the process.

### Usage
#### Using the Executable
To use the pre-built executable, download `msg_eml_converter.exe` from the [latest release](https://github.com/TohokuSteelKiki/MsgToEmlConverter/releases/latest) and run it. Follow the prompts to convert your MSG files to EML.

#### Creating the Executable using PyInstaller
If you prefer to create the executable yourself, follow these steps:
1. Install PyInstaller:
   ~~~sh
   pip install pyinstaller
   ~~~
2. Navigate to the directory containing the `msg_eml_converter.py` file and run PyInstaller:
   ~~~sh
   pyinstaller --onefile msg_eml_converter.py
   ~~~
3. The executable will be created in the `dist` directory.

#### Running the Program from Source
1. Add Remote Repository:
   ~~~sh
   git remote add origin https://github.com/TohokuSteelKiki/MsgToEmlConverter.git
   ~~~
2. Install Dependencies:
   Ensure you have the required Python libraries installed:
   ~~~sh
   pip install extract-msg
   ~~~
3. Run the Program:
   Execute the program and follow the prompts to convert your MSG files to EML:
   ~~~sh
   python msg_eml_converter.py
   ~~~
### Requirements
- Python 3.6+
- extract-msg library

### License
This project is licensed under the MIT License.

---

## :jp: 日本語版

### 説明
このPythonプログラムは、Microsoft OutlookのMSGファイルをEML形式に変換し、添付ファイルとメタデータを保持します。
指定されたディレクトリ内の複数のMSGファイルを処理し、EML形式に変換して、別のディレクトリに変換されたファイルを保存するように設計されています。さらに、受信日時をメール件名の先頭に追加するオプションを提供し、整理された時系列のメール記録を保証します。

### 特徴
- バッチ変換: ディレクトリ内の複数のMSGファイルをEML形式に変換します。
- 添付ファイル処理: MSGファイルからすべての添付ファイルを抽出し、変換されたEMLファイルに含めます。
- メタデータ保持: 送信者、受信者、件名、日付などの重要なメールメタデータを保持します。
- 日付プレフィックスオプション: 受信日時をメール件名の先頭に追加するオプションを提供し、より良い整理を実現します。
- 一時ファイル管理: 変換中に一時ファイルを管理し、プロセス終了後にクリーンアップします。

### 使用方法
#### 実行可能ファイルの使用
事前にビルドされた実行可能ファイルを使用するには、[最新のリリース](https://github.com/TohokuSteelKiki/MsgToEmlConverter/releases/latest)から`msg_eml_converter.exe`をダウンロードして実行します。プロンプトに従ってMSGファイルをEMLに変換します。

#### PyInstallerを使用して実行可能ファイルを作成する
自分で実行可能ファイルを作成する場合、以下の手順に従ってください:
1. PyInstallerをインストールします:
   ~~~sh
   pip install pyinstaller
   ~~~
2. `msg_eml_converter.py`ファイルが含まれているディレクトリに移動し、PyInstallerを実行します:
   ~~~sh
   pyinstaller --onefile msg_eml_converter.py
   ~~~
3. 実行可能ファイルは`dist`ディレクトリに作成されます。

#### ソースからプログラムを実行する
1. リモートリポジトリの追加:
   ~~~sh
   git remote add origin https://github.com/TohokuSteelKiki/MsgToEmlConverter.git
   ~~~
2. 依存関係のインストール:
   必要なPythonライブラリをインストールします:
   ~~~sh
   pip install extract-msg
   ~~~
3. プログラムの実行:
   プログラムを実行し、プロンプトに従ってMSGファイルをEMLに変換します:
   ~~~sh
   python msg_eml_converter.py
   ~~~
### 必要条件
- Python 3.6以上
- extract-msg ライブラリ

### ライセンス
このプロジェクトはMITライセンスの下でライセンスされています。
