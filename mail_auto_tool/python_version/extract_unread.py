import win32com.client as win32
import pandas as pd
from datetime import datetime

# ==============================
# 設定
# ==============================
OUTPUT_PATH = "mail_data.xlsx"  # 出力ファイル名
UNREAD_ONLY = True               # 未読のみ抽出するか

# ==============================
# 未読メール抽出
# ==============================
def extract_unread_emails():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = 受信トレイ

    mails = inbox.Items
    mails = mails.Restrict("[Unread]=True") if UNREAD_ONLY else mails
    mails.Sort("[ReceivedTime]", True)

    data = []
    for mail in mails:
        try:
            if mail.Class == 43:  # メールアイテムのみ
                data.append({
                    "ReceivedTime": mail.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                    "SenderName": mail.SenderName,
                    "Subject": mail.Subject,
                    "Body": mail.Body[:500],  # 長文防止で500文字まで
                    "To": "",  # 後で追記する用
                    "SendFlag": 0  # 送信用フラグ
                })
        except Exception as e:
            print(f"スキップ: {e}")

    if not data:
        print("📭 未読メールはありません。")
        return

    df = pd.DataFrame(data)
    df.to_excel(OUTPUT_PATH, index=False)
    print(f"✅ {len(df)}件の未読メールを抽出しました → {OUTPUT_PATH}")

# ==============================
# メイン
# ==============================
if __name__ == "__main__":
    extract_unread_emails()
