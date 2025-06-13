# 請求書自動生成ツール（GUI対応）

このツールは、テンプレートExcelファイルとCSVデータを元に、複数の請求書・見積書・納品書を一括生成するPythonアプリケーションです。
GUIにより操作が直感的で、非エンジニアでも利用可能です。

---

## 🔧 主な機能

* 顧客情報CSVをもとに、Excelテンプレートへデータを差し込み出力
* 請求書／見積書／納品書の3種テンプレートに対応
* GUIでCSVファイルとテンプレート種別を選択可能
* Excelアプリ未インストールでも動作（openpyxl使用）

---

## 🚀 使い方

1. 必要なライブラリをインストール：

   ```bash
   pip install openpyxl
   ```

2. 仮想環境を有効化（任意）：

   ```bash
   .\excel-env\Scripts\activate
   ```

3. GUIを起動：

   ```bash
   python auto_invoice_gui.py
   ```

4. GUIでCSVとテンプレートを選択 → \[出力実行] ボタンをクリック

---

## 📁 ファイル構成

```
excel_automation_project/
├── auto_invoice_gui.py         # GUI付き本体スクリプト
├── generate_templates.py       # テンプレート作成用スクリプト
├── client_data.csv             # 顧客データCSV
├── template_invoice.xlsx       # 請求書テンプレート
├── template_estimate.xlsx      # 見積書テンプレート
├── template_delivery.xlsx      # 納品書テンプレート
└── .gitignore                  # 仮想環境など除外設定
```

---

## 💡 発展例

* **PDF自動変換対応**：xlsx出力後、`xlsx2pdf` や `win32com.client` を用いてPDF出力対応
* **メール一括送信**：`smtplib` や `yagmail` を用いて、各請求書をメール送信可能に
* **Webフォーム連携**：FlaskやStreamlitを使って、Web UIからCSVアップロードと出力処理を可能に
* **Googleスプレッドシート対応**：Google APIと連携し、CSVの代わりにスプレッドシートを読み込む機能

---


## 📜 ライセンス

MIT License
