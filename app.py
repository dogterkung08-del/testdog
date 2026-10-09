import os
import re
import pandas as pd
from flask import Flask, request, abort

from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage
)

app = Flask(__name__)

# ดึง Key จาก Environment Variables หรือใช้ค่า Token สำรองหากหาไม่เจอ
CHANNEL_ACCESS_TOKEN = os.environ.get(
    'CHANNEL_ACCESS_TOKEN', 
    'XVft4ru1WrRaNg3p3VDPzNOjy3gmLuW1hdQuwh/qOKPjqr5hTL3UfIcZzFUE98wCyB713E3gurGNGTNqVf7wzmEjsiXZZ9Af7HUqM1I8K+DLycAHYDIgjQoIrkxJ1DLUNnlI8yeTq6u6w71N0VYVEAdB04t89/1O/w1cDnyilFU='
)
CHANNEL_SECRET = os.environ.get(
    'CHANNEL_SECRET', 
    'e3985f1f44458497d3d65015caa1f9d1'
)

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

EXCEL_FILE = "data.xlsx"
DIRECTOR_FILE = "name.xlsx"

def search_data(user_text):
    try:
        # 1. ค้นหาข้อมูล ผอ. จากไฟล์ name.xlsx
        if "ผอ" in user_text:
            search_term = re.sub(r"ผอ\.|ผอ|ของ", "", user_text).strip()
            if not search_term:
                search_term = user_text

            if not os.path.exists(DIRECTOR_FILE):
                return f"⚠️ ไม่พบไฟล์ {DIRECTOR_FILE} ในระบบ กรุณาตรวจสอบชื่อไฟล์ครับ"

            df_dir = pd.read_excel(DIRECTOR_FILE, dtype=str).fillna("")
            mask = df_dir.iloc[:, 0].str.contains(
                search_term, regex=False, case=False, na=False
            ) | df_dir.iloc[:, 1].str.contains(
                search_term, regex=False, case=False, na=False
            )
            matched_rows = df_dir[mask]

            if not matched_rows.empty:
                results_text = f"📋 ผลการค้นหาข้อมูล ผอ. ({search_term}):\n\n"
                for _, row in matched_rows.iterrows():
                    org_name = row.iloc[0] if len(row) > 0 else ""
                    director_name = row.iloc[1] if len(row) > 1 else ""
                    updated_date = row.iloc[2] if len(row) > 2 else ""

                    results_text += f"🏢 {org_name}\n"
                    results_text += f"👤 ผอ.: {director_name}\n"
                    if updated_date:
                        results_text += f"📅 อัปเดตเมื่อ: {updated_date}\n"
                    results_text += "------------------------------\n"

                return results_text.strip()
            else:
                return "พิมพ์อะไรผิดไปรึเปล่า ลองดูใหม่ดิ วุ้วววววว!!"

        # 2. ค้นหาข้อมูลทั่วไปจากไฟล์ data.xlsx
        if not os.path.exists(EXCEL_FILE):
            return f"⚠️ ไม่พบไฟล์ {EXCEL_FILE} ในระบบ กรุณาตรวจสอบชื่อไฟล์ครับ"

        df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
        mask = df.astype(str).apply(
            lambda x: x.str.contains(user_text, regex=False, case=False, na=False)
        ).any(axis=1)
        matched_rows = df[mask]

        if not matched_rows.empty:
            response_text = ""
            for _, row_data in matched_rows.head(10).iterrows():
                val_a = row_data.iloc[0] if len(row_data) > 0 else ""
                val_b = row_data.iloc[1] if len(row_data) > 1 else ""
                val_c = row_data.iloc[2] if len(row_data) > 2 else ""
                val_d = row_data.iloc[3] if len(row_data) > 3 else ""

                response_text += f"{val_a}\n{val_b}\n{val_c}\n{val_d}\n\n"

            if len(matched_rows) > 10:
                response_text += f"\n*(แสดง 10 รายการแรก จากทั้งหมด {len(matched_rows)} รายการ)*"

            return response_text.strip()
        else:
            return "พิมพ์อะไรผิดไปรึเปล่า ลองดูใหม่ดิ วุ้วววววว!!"

    except Exception:
        return "อิหยัง ？"

# Route สำหรับเช็กสถานะบอท (ใช้ป้องกัน Server หลับ)
@app.route("/", methods=['GET'])
def health_check():
    return "Bot is running fine!", 200

# Route สำหรับรับ Webhook จาก LINE
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

# ทำงานเมื่อมีข้อความเข้าในกลุ่มหรือแชตส่วนตัว
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_text = event.message.text.strip()
    reply_text = search_data(user_text)

    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=reply_text)
    )

if __name__ == "__main__":
    app.run(port=5000)
