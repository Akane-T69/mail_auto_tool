<h1 align="center">📧 Excel × Outlook 自動化ツール</h1>

<p align="center">
  <img src="https://img.shields.io/badge/Excel-VBA-green?logo=microsoft-excel&logoColor=white" />
  <img src="https://img.shields.io/badge/Outlook-Automation-blue?logo=microsoft-outlook&logoColor=white" />
  <img src="https://img.shields.io/badge/License-Free-lightgrey" />
  <img src="https://img.shields.io/badge/Build-Stable-brightgreen" />
</p>

---

### 💡 概要
このツールは、**Outlookの未読メールをExcelに自動抽出**し、  
さらにExcel上のデータをもとに**自動送信メールを生成・送信**する自動化システムです。  
日常の監視・報告・顧客対応メールの効率化を目的としています。

---

### 🧩 主な機能
- 📥 **未読メール抽出機能**：日時・差出人・件名・本文をExcelに自動取得  
- 📤 **自動送信機能**：Excelの内容をもとに複数宛先へ一括メール送信  
- 🧾 **テンプレート対応**：「template_mail.xlsx」から本文テンプレートを読み込み  
- ⚙️ **運用自動化向け構成**：監視・報告業務の省力化を想定した設計  

---

### 🧠 想定ユースケース
- サーバー監視報告や障害通知を自動メール化  
- 顧客・社内向け定期レポートの自動配信  
- 日報・進捗報告などのExcelベース自動送信  

---

### 🛠 技術スタック
| カテゴリ | 使用技術 |
|:--|:--|
| 言語 | VBA (Visual Basic for Applications) |
| アプリ連携 | Microsoft Outlook / Excel |
| OS環境 | Windows 10 / 11 |
| バージョン | Office 2016以上推奨 |

---

### 📂 フォルダ構成
mail_auto_tool/
├─ main.vba ← メインスクリプト（Excelモジュール用）
├─ template_mail.xlsx ← メールテンプレート（A1～C15に文面）
├─ sample_data.xlsx ← ダミーデータ（宛先・件名・本文）
└─ README.md ← 本ファイル

---



---

### 🚀 実行方法
1. Excelの開発タブからマクロを有効化  
2. 「ExtractUnreadEmails」マクロを実行 → Outlook未読メールを抽出  
3. 「SendMailFromExcel」マクロを実行 → メールを自動送信  

---

### 📜 ライセンス
このプロジェクトは個人ポートフォリオ用に作成されたもので、自由に閲覧・参考利用可能です。

---

### 👤 作者
**Akane**  
職種：サーバー監視・保守業務  
興味分野：業務効率化・自動化・クラウド運用  

💬 キーワード

#ExcelVBA #OutlookAutomation #メール自動化 #ポートフォリオ #業務効率化

© 2025 Akane


✅ 注意事項

実行にはOutlookとExcelの両方がインストールされている必要があります
VBA実行時はマクロを有効にしてください
実運用前にテストメールで動作確認を行ってください
実際の個人情報や顧客データでの動作は控えてください（ポートフォリオ用）


💼 想定利用シーン

監視業務の「障害検知報告」や「日次レポート送信」
顧客対応の「定型返信メール」自動化
チーム間共有の「未読メール一覧抽出」など


🗣 作者コメント

本ツールは、手作業の報告・メール対応を自動化し、
「人が判断すべき業務に時間を使える環境をつくる」ことを目的に開発しました。
今後はログ監視ツールやクラウド通知への展開も視野に入れています。
