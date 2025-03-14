# つくば市ゴミ収集カレンダー作成ツール

![ゴミ収集カレンダー](https://img.shields.io/badge/つくば市-ゴミ収集カレンダー-brightgreen)

## 📝 概要

このアプリケーションは、つくば市のゴミ収集日程をExcelファイルから読み込み、iCalendar形式のカレンダーファイルを作成するツールです。作成したiCalendarファイルは、Googleカレンダー、Outlookなどのカレンダーアプリに簡単にインポートできます。

## ✨ 特徴

- 📊 Excelファイルからゴミ収集日程を読み込み
- 📆 iCalendar形式のカレンダーファイルを作成
- 🏙️ 地区ごとのカレンダー作成に対応
- 🗑️ 複数のゴミ種類に対応（燃やせるごみ、びん、ペットボトルなど）
- 🖥️ シンプルで使いやすいGUIインターフェース

## 🔧 インストール方法

### 必要条件

- Python 3.12以上
- 必要なパッケージ：
  - openpyxl
  - icalendar
  - pytz

### インストール手順

1. リポジトリをクローンまたはダウンロードします。
2. 仮想環境を作成し、アクティベートします：

```bash
python -m venv venv
# Windowsの場合
venv\Scripts\activate
# macOS/Linuxの場合
source venv/bin/activate
```

3. 必要なパッケージをインストールします：

```bash
pip install -r requirements.txt
```

## 🚀 使い方

1. アプリケーションを起動します：

```bash
python CreateGomiCalender.py
```

2. GUIが表示されたら、以下の手順で操作します：

   ![操作手順](https://via.placeholder.com/600x400?text=操作手順イメージ)

   1. **「Select Excel Directory」ボタン**をクリックして、Excelファイルが保存されているディレクトリを選択します。
   2. ドロップダウンリストからExcelファイルを選択し、**「Select Excel」ボタン**をクリックします。
   3. ドロップダウンリストから地区名を選択し、**「Select Region」ボタン**をクリックします（デフォルトは「並木」）。
   4. **「Create iCal」ボタン**をクリックして、iCalendarファイルを作成します。

3. 作成されたiCalendarファイルは、`ical`ディレクトリに保存されます。ファイル名は「YYYYMM_地区名.ics」の形式です（例：202503_並木.ics）。

## 📁 ファイル形式

### Excelファイル

- ファイル名：YYYYMM_calendar.xlsx（例：202503_calendar.xlsx）
- 1行目：ヘッダー行
- 2行目以降：地区ごとのデータ
- 列構造：
  - 1列目：地区名
  - 2列目：備考
  - 3列目：燃やせるごみの収集日
  - 4列目：びんの収集日
  - 5列目：スプレー容器の収集日
  - 6列目：ペットボトルの収集日
  - 7列目：燃やせないごみの収集日
  - 8列目：古紙・古布の収集日
  - 9列目：プラスチック製容器包装の収集日
  - 10列目：かんの収集日
  - 11列目：粗大ごみ（予約制）の収集日

### iCalendarファイル

- ファイル名：YYYYMM_地区名.ics（例：202503_並木.ics）
- 形式：標準的なiCalendar形式（RFC 5545準拠）
- 内容：各ゴミ種類の収集日がイベントとして登録されています

## 📱 カレンダーアプリへのインポート方法

### Googleカレンダー

1. [Googleカレンダー](https://calendar.google.com/)にアクセスします。
2. 画面右上の歯車アイコン（設定）をクリックし、「設定」を選択します。
3. 左側のメニューから「インポート/エクスポート」を選択します。
4. 「ファイルを選択」をクリックし、作成したiCalendarファイル（.ics）を選択します。
5. カレンダーを選択し、「インポート」をクリックします。

### iPhoneカレンダー

1. 作成したiCalendarファイル（.ics）をメールで自分に送信します。
2. iPhoneでメールを開き、添付ファイル（.ics）をタップします。
3. 「カレンダーに追加」をタップします。

### Outlookカレンダー

1. Outlookを開きます。
2. 「ファイル」タブをクリックし、「開く」→「インポート」を選択します。
3. 「iCalendarまたはvCalendarファイルのインポート」を選択し、「次へ」をクリックします。
4. 作成したiCalendarファイル（.ics）を選択し、「OK」をクリックします。

## ❓ よくある質問

### Q: 特定の地区のデータがありません
A: Excelファイルに該当地区のデータが含まれているか確認してください。地区名の表記が正確であることを確認してください。

### Q: iCalendarファイルが作成されません
A: 以下を確認してください：
- Excelファイルが正しい形式であるか
- 選択した地区のデータが存在するか
- 必要なパッケージがすべてインストールされているか

### Q: 特定のゴミ種類の収集日が表示されません
A: Excelファイルの該当列に正しい形式で日付が入力されているか確認してください。日付は「YYYY/MM/DD」形式である必要があります。

## 🔄 更新履歴

- 2025/03/14: 初版リリース
  - 燃やせるごみの日が正しく反映されない問題を修正
  - Excelファイルのヘッダー行との表記の統一

## 📄 ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は[LICENSE](LICENSE)ファイルを参照してください。

## 👥 貢献

バグ報告や機能リクエストは、Issueを作成してください。プルリクエストも歓迎します！

---

作成：2025年3月