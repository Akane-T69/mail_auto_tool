# 📬 メール自動抽出・送信ツール（Outlook × Excel × Python）

![Python](https://img.shields.io/badge/Python-3.11+-blue?logo=python&logoColor=white)
![Outlook](https://img.shields.io/badge/Outlook-Automation-blue?logo=microsoft-outlook)
![Excel](https://img.shields.io/badge/Excel-openpyxl-green?logo=microsoft-excel)
![Status](https://img.shields.io/badge/Build-Stable-brightgreen)

## 🧩 概要
Outlookから未読メールを自動抽出し、Excelに記録。
Excel上で「送信フラグ」を立てると、テンプレートに基づいて自動送信します。

## ⚙️ 構成
mail_auto_tool/
├─ extract_unread.py    ← Outlookの未読メールをExcelに抽出
├─ send_mail.py         ← Excelからメール送信（テンプレート使用＋ログ出力）
├─ mail_data.xlsx       ← 未読抽出で生成される or 編集して送信用に使う
├─ template_mail.xlsx   ← メール本文テンプレート
├─ send_log.xlsx        ← 送信結果ログ（自動生成）
└─ README.md

## 🚀 使い方
1. Outlookが起動できる状態で `main.py` を実行  
2. `mail_data.xlsx` に未読メールが出力されます  
3. 「送信フラグ」列に `1` を入力した行を対象に再実行で送信  
4. テンプレートは `{body}` 部分にメール本文を差し込みます

## ⚠️ 注意事項
- Windows + Outlook 環境が必須  
- 自動送信前に内容を確認してください  
- 機密情報・個人メールでの利用は禁止

## 💬 作者コメント
VBAで作成した業務自動化ツールをPython化。
実務でのExcel運用・監視業務を効率化する実用コードです。
