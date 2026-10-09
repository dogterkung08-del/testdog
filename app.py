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

# ดึง Key จาก Environment Variables หรือใช้ค่า Token สำรอง
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

def search_excel_all_sheets(file_path, search_term):
    """ ฟังก์ชันช่วยค้นหาคำจากทุก Sheet ในไฟล์ Excel แบบ Contains """
    if not os.path.exists(file_path):
        return None, f"⚠️ ไม่พบไฟล์ {file_path} ในระบบ กรุณาตรวจสอบชื่อไฟล์ครับ"

    xls = pd.ExcelFile(file_path)
    all_matched = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")
        # ค้นหาทุกคอลัมน์ที่มีคำค้นหาปนอยู่ (ไม่สนตัวพิมพ์เล็ก-ใหญ่ และค้นหาบางส่วนได้)
        mask = df.astype(str).apply(
            lambda x: x.str.contains(search_term, regex=False, case=False, na=False)
        ).any(axis=1)
        matched = df[mask]
        if not matched.empty:
            all_matched.append(matched)

    if all_matched:
        return pd.concat(all_matched, ignore_index=True), None
    return pd.DataFrame(), None

def search_data(user_text):
    try:
        # ยุบช่องว่างที่เว้นวรรคซ้ำๆ ให้เหลือ 1 ช่องว่าง และตัดช่องว่างหน้า-หลัง
        cleaned_text = re.sub(r"\s+", " ", user_text).strip()

        # -------------------------------------------------------------
        # 1. ค้นหาข้อมูล ผอ. จากไฟล์ name.xlsx
        # -------------------------------------------------------------
        if "ผอ" in cleaned_text:
            # ตัดคำว่า "ผอ.", "ผอ", "ของ" ออก และจัดการช่องว่าง
            search_term = re.sub(r"ผอ\.|ผอ|ของ", " ", cleaned_text)
            search_term = re.sub(r"\s+", " ", search_term).strip()

            if not search_term:
                return "กรุณาระบุชื่อหน่วยงานหรือจังหวัด เช่น 'ผอ. กระบี่' ครับ"

            matched_df, err = search_excel_all_sheets(DIRECTOR_FILE, search_term)
            if err:
                return err

            if not matched_df.empty:
                results_text = f"📋 ผลการค้นหาข้อมูล ผอ. ({search_term}):\n\n"
                for _, row in matched_df.iterrows():
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

        # -------------------------------------------------------------
        # 2. ค้นหาข้อมูลทั่วไปจากไฟล์ data.xlsx
        # -------------------------------------------------------------
        search_term = cleaned_text
        matched_df, err = search_excel_all_sheets(EXCEL_FILE, search_term)
        if err:
            return err

        if not matched_df.empty:
            response_text = ""
            for _, row_data in matched_df.head(10).iterrows():
                val_a = row_data.iloc[0] if len(row_data) > 0 else ""
                val_b = row_data.iloc[1] if len(row_data) > 1 else ""
                val_c = row_data.iloc[2] if len(row_data) > 2 else ""
                val_d = row_data.iloc[3] if len(row_data) > 3 else ""

                response_text += f"{val_a}\n{val_b}\n{val_c}\n{val_d}\n\n"

            if len(matched_df) > 10:
                response_text += f"\n*(แสดง 10 รายการแรก จากทั้งหมด {len(matched_df)} รายการ)*"

            return response_text.strip()

        return "พิมพ์อะไรผิดไปรึเปล่า ลองดูใหม่ดิ วุ้วววววว!!"

    except Exception:
        return "อิหยัง ？"

# Route สำหรับเช็กสถานะบอท
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

# ทำงานเมื่อมีข้อความเข้า
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
