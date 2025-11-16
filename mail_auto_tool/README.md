# Outlookメール抽出＆自動返信ツール（Excel連携）

## 概要
Outlookの未読メールをExcelに抽出し、指定セルにフラグが立った行の情報を使って自動返信メールを送信するツールです。  
日常的な問い合わせ対応や報告業務の自動化に活用できます。

## フォルダ構成
- vba_version/
  - outlook_to_excel.xlsm : 未読メール抽出用VBA
  - auto_mail_sender.xlsm : 自動返信VBA
  - template_mail.xlsx : メールテンプレート
- python_version/ : Python版移行用
- sample_result.xlsx : 実行結果サンプル

## 動作環境
- Windows 10/11
- Microsoft Excel 2016以降
- Microsoft Outlook 2016以降

## 使い方
1. Excelマクロを有効化
2. `outlook_to_excel` を実行して未読メールを抽出
3. Excelのフラグ列に「1」を入力
4. `auto_mail_sender` を実行してメールを送信

## スキルアピールポイント
- Excel・Outlookの自動化（VBA）
- 条件トリガーでの処理制御
- 業務改善・運用効率化の実践例
