import os
import re
from datetime import datetime, timezone, timedelta
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from flask import Flask, request, abort, send_from_directory

from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage, ImageSendMessage
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

# กำหนดชื่อไฟล์ Excel และระบบ
EXCEL_FILE = "data.xlsx"
DIRECTOR_FILE = "name.xlsx"
LOCATION_FILE = "lo.xlsx"
BG_FILENAME = "background.png"

IMG_W = 996
IMG_H = 1242

MONTHS_TH = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

BANK_IMAGE_FILES = {
    "ธ.กสิกรไทย": "ธ.กสิกรไทย.png",
    "ธ.กรุงเทพ": "ธ.กรุงเทพ.png",
    "ธ.กรุงไทย": "ธ.กรุงไทย.png",
    "ธ.กรุงศรีอยุธยา": "ธ.กรุงศรีอยุธยา.png",
    "ธ.เกียรตินาคินภัทร": "ธ.เกียรตินาคินภัทร.png",
    "ธ.ซีไอเอ็มบี": "ธ.ซีไอเอ็มบี.png",
    "ธ.ไทยพาณิชย์": "ธ.ไทยพาณิชย์.png",
    "ธ.ก.ส.": "ธ.เพื่อการเกษตรและสหกรณ์.png",
    "ธ.ทหารไทยธนชาต": "ธ.ทหารไทยธนชาต.png",
    "ธ.ยูโอบี": "ธ.ยูโอบี.png",
    "ธ.ออมสิน": "ธ.ออมสิน.png",
    "ธ.อาคารสงเคราะห์": "ธ.อาคารสงเคราะห์.png"
}

FONT_WEIGHTS = {
    "Light": "KuriousLoopedCond-Light.ttf",
    "Regular": "KuriousLoopedCond-Regular.ttf",
    "Medium": "KuriousLoopedCond-Medium.ttf",
    "SemiBold": "KuriousLoopedCond-SemiBold.ttf"
}

FIELDS_CONFIG = {
    "title": {"x": 67, "y": 21, "size": 47, "align": "left", "weight": "SemiBold", "color": "#333333", "opacity": 96, "default": "โอนเงินสำเร็จ"},
    "datetime": {"x": 67, "y": 98, "size": 39, "align": "left", "weight": "Regular", "color": "#555555", "opacity": 96},
    "sender_name": {"x": 243, "y": 235, "size": 39, "align": "left", "weight": "Regular", "color": "#222222", "opacity": 96, "default": "นาย สุริยา มาสุข"},
    "sender_bank": {"x": 243, "y": 298, "size": 39, "align": "left", "weight": "Medium", "color": "#555555", "opacity": 96, "default": "ธ.กสิกรไทย"},
    "sender_acc": {"x": 243, "y": 362, "size": 39, "align": "left", "weight": "Medium", "color": "#555555", "opacity": 96, "default": "5743"},
    "receiver_name": {"x": 243, "y": 550, "size": 39, "align": "left", "weight": "Regular", "color": "#222222", "opacity": 96, "default": "นางสาว ไพลิน ไตรมงคล"},
    "receiver_bank": {"x": 243, "y": 612, "size": 39, "align": "left", "weight": "Medium", "color": "#555555", "opacity": 96, "default": "ธ.กสิกรไทย"},
    "receiver_acc": {"x": 243, "y": 676, "size": 39, "align": "left", "weight": "Medium", "color": "#555555", "opacity": 96, "default": "2403"},
    "ref_no": {"x": 605, "y": 870, "size": 37, "align": "right", "weight": "Medium", "color": "#444444", "opacity": 96, "default": "015334170101BOR09546"},
    "amount": {"x": 605, "y": 992, "size": 39, "align": "right", "weight": "Regular", "color": "#222222", "opacity": 96, "default": "25,000.00"},
    "fee": {"x": 605, "y": 1118, "size": 39, "align": "right", "weight": "Regular", "color": "#222222", "opacity": 96, "default": "0.00"},
    "label_scan": {"x": 813, "y": 1130, "size": 32, "align": "center", "weight": "Light", "color": "#555555", "opacity": 96, "default": "สแกนตรวจสอบสลิป"}
}

# ฟังก์ชันจัดการข้อความและสลิป
def format_short_sender_name(full_name):
    parts = full_name.strip().split()
    if len(parts) >= 2:
        parts[-1] = parts[-1][0]
        return " ".join(parts)
    return full_name

def format_currency(value_str):
    clean_val = ''.join([c for c in str(value_str) if c.isdigit() or c == '.'])
    if not clean_val:
        return "0.00"
    try:
        num = float(clean_val)
        return "{:,.2f}".format(num)
    except ValueError:
        return value_str

def update_ref_no_time(ref_no, time_str):
    digits = ''.join(filter(str.isdigit, time_str)).zfill(4)[:4]
    if len(ref_no) >= 10:
        return ref_no[:6] + digits + ref_no[10:]
    return ref_no

def hex_to_rgba(hex_color, opacity_percent):
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16) if len(hex_color) >= 6 else 0
    g = int(hex_color[2:4], 16) if len(hex_color) >= 6 else 0
    b = int(hex_color[4:6], 16) if len(hex_color) >= 6 else 0
    alpha = int((max(0, min(100, opacity_percent)) / 100.0) * 255)
    return (r, g, b, alpha)

def get_font(weight_name, size):
    font_filename = FONT_WEIGHTS.get(weight_name, "KuriousLoopedCond-Regular.ttf")
    try:
        return ImageFont.truetype(font_filename, size)
    except IOError:
        try:
            return ImageFont.truetype("tahoma.ttf", size)
        except IOError:
            return ImageFont.load_default()

def generate_default_slip():
    """สร้างรูปภาพสลิปตามค่าเริ่มต้น (รองรับเวลาปัจจุบันประเทศไทย)"""
    if os.path.exists(BG_FILENAME):
        img = Image.open(BG_FILENAME).convert("RGBA").resize((IMG_W, IMG_H), Image.Resampling.LANCZOS)
    else:
        img = Image.new("RGBA", (IMG_W, IMG_H), (234, 234, 234, 255))

    draw = ImageDraw.Draw(img)

    bank_name = "ธ.กสิกรไทย"
    overlay_filename = BANK_IMAGE_FILES.get(bank_name, "ธ.กสิกรไทย.png")
    if os.path.exists(overlay_filename):
        try:
            overlay_img = Image.open(overlay_filename).convert("RGBA")
            target_w = 207
            w, h = overlay_img.size
            target_h = int(h * (target_w / w))
            overlay_img = overlay_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
            img.paste(overlay_img, (13, 527), overlay_img)
        except Exception:
            pass

    # ใช้เวลาปัจจุบันของประเทศไทย
    tz_th = timezone(timedelta(hours=7))
    now_th = datetime.now(tz_th)
    day_str = str(now_th.day)
    month_str = MONTHS_TH[now_th.month - 1]
    year_str = str(now_th.year + 543)[-2:]
    time_str = now_th.strftime("%H:%M")

    dt_cfg = FIELDS_CONFIG["datetime"]
    dt_text = f"{day_str} {month_str} {year_str}  {time_str} น."
    dt_font = get_font(dt_cfg["weight"], dt_cfg["size"])
    draw.text((dt_cfg["x"], dt_cfg["y"]), dt_text, fill=hex_to_rgba(dt_cfg["color"], dt_cfg["opacity"]), font=dt_font)

    sender_display_name = format_short_sender_name("นาย สุริยา มาสุข")
    amt_fmt = format_currency("25000")
    fee_fmt = format_currency("0")
    ref_updated = update_ref_no_time("015334170101BOR09546", time_str)

    values = {
        "title": "โอนเงินสำเร็จ",
        "sender_name": sender_display_name,
        "sender_bank": "ธ.กสิกรไทย",
        "sender_acc": "xxx-x-x5743-x",
        "receiver_name": "นางสาว ไพลิน ไตรมงคล",
        "receiver_bank": bank_name,
        "receiver_acc": "xxx-x-x2403-x",
        "ref_no": ref_updated,
        "amount": f"{amt_fmt} บาท",
        "fee": f"{fee_fmt} บาท",
        "label_scan": "สแกนตรวจสอบสลิป"
    }

    for key, text_val in values.items():
        if key == "datetime":
            continue
        cfg = FIELDS_CONFIG.get(key)
        if not cfg:
            continue

        font = get_font(cfg["weight"], cfg["size"])
        color_rgba = hex_to_rgba(cfg["color"], cfg["opacity"])

        if cfg["align"] == "right":
            bbox = draw.textbbox((0, 0), text_val, font=font)
            draw_x = cfg["x"] - (bbox[2] - bbox[0])
        elif cfg["align"] == "center":
            bbox = draw.textbbox((0, 0), text_val, font=font)
            draw_x = cfg["x"] - ((bbox[2] - bbox[0]) / 2)
        else:
            draw_x = cfg["x"]

        draw.text((draw_x, cfg["y"]), text_val, fill=color_rgba, font=font)

    output_path = "generated_slip.png"
    img.convert("RGB").save(output_path, quality=95)
    return output_path

# ฟังก์ชันค้นหา Excel ทุก Sheet
def search_excel_all_sheets(file_path, search_term):
    if not os.path.exists(file_path):
        return None, f"⚠️ ไม่พบไฟล์ {file_path} ในระบบ กรุณาตรวจสอบชื่อไฟล์ครับ"

    xls = pd.ExcelFile(file_path)
    all_matched = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")
        mask = df.astype(str).apply(
            lambda x: x.str.contains(search_term, regex=False, case=False, na=False)
        ).any(axis=1)
        matched = df[mask]
        if not matched.empty:
            all_matched.append(matched)

    if all_matched:
        return pd.concat(all_matched, ignore_index=True), None
    return pd.DataFrame(), None

def search_data(user_text, base_url):
    try:
        cleaned_text = re.sub(r"\s+", " ", user_text).strip()

        # 1. ตรวจสอบคำสั่งเรียกสลิป
        if cleaned_text in ["/สลิป", "สลิป"]:
            generate_default_slip()
            image_url = f"{base_url.rstrip('/')}/image/generated_slip.png"
            return {
                "type": "image",
                "original_url": image_url,
                "preview_url": image_url,
                "text": "✅ สลิปการโอนเงิน ธ.กสิกรไทย ของคุณพร้อมแล้วครับ!"
            }

        # 2. ตรวจสอบคำสั่งช่วยเหลือ /help
        if cleaned_text in ["/help", "help", "ช่วยเหลือ"]:
            help_text = (
                "🤖 คู่มือการใช้งานและคำสั่งทั้งหมดของบอท:\n"
                " • /สลิป - สร้างสลิปโอนเงิน ธ.กสิกรไทย\n"
                " • /help - แสดงหน้าจอคู่มือคำสั่งทั้งหมดนี้\n"
                " ระบบค้นหาข้อมูล กรมพินิจ \n"
                " • ค้นหาข้อมูลทั่วไป: พิมพ์ชื่อจังหวัดหรือคำค้นหา\n"
                " • ค้นหาข้อมูล ผอ.: พิมพ์คำว่า 'ผอ. ชื่อจังหวัด' (เช่น ผอ. นนทบุรี)\n"
                " • ค้นหาโลเคชั่น/แผนที่: พิมพ์ 'โลเคชั่น ชื่อจังหวัด' หรือ 'โล ชื่อจังหวัด'"
            )
            return {"type": "text", "text": help_text}

        # 3. ค้นหาข้อมูลโลเคชั่น / แผนที่ จากไฟล์ lo.xlsx
        is_location_query = (
            re.search(r"โลเคชั่น|โลเคชัน|location|map|แผนที่|ขอโล", cleaned_text, re.IGNORECASE) 
            or (re.search(r"(^โล|โล$)", cleaned_text) and cleaned_text != "พิษณุโลก")
        )

        if is_location_query:
            search_term = re.sub(r"ขอ|โลเคชั่น|โลเคชัน|location|map|แผนที่", " ", cleaned_text, flags=re.IGNORECASE)
            search_term = re.sub(r"^โล|โล$", " ", search_term)
            search_term = re.sub(r"\s+", " ", search_term).strip()

            if not search_term:
                return {"type": "text", "text": "กรุณาระบุชื่อจังหวัด เช่น 'โลเคชั่น นนทบุรี' ครับ"}

            matched_df, err = search_excel_all_sheets(LOCATION_FILE, search_term)
            if err:
                return {"type": "text", "text": err}

            if not matched_df.empty:
                results_text = f"📍 ผลการค้นหาโลเคชั่น ({search_term}):\n\n"
                for _, row in matched_df.iterrows():
                    org_name = row.iloc[0] if len(row) > 0 else ""
                    map_url = row.iloc[1] if len(row) > 1 else ""
                    updated_date = row.iloc[2] if len(row) > 2 else ""

                    results_text += f"🏢 {org_name}\n"
                    results_text += f"📍 โลเคชั่น: {map_url}\n"
                    if updated_date:
                        results_text += f"📅 อัปเดตเมื่อ: {updated_date}\n"
                    results_text += "------------------------------\n"
                return {"type": "text", "text": results_text.strip()}
            else:
                return {"type": "text", "text": "พิมพ์อะไรผิดไปรึเปล่า ลองดูใหม่ดิ วุ้วววววว!!"}

        # 4. ค้นหาข้อมูล ผอ. จากไฟล์ name.xlsx
        if "ผอ" in cleaned_text:
            search_term = re.sub(r"ผอ\.|ผอ|of|ของ", " ", cleaned_text)
            search_term = re.sub(r"\s+", " ", search_term).strip()

            if not search_term:
                return {"type": "text", "text": "กรุณาระบุชื่อหน่วยงานหรือจังหวัด เช่น 'ผอ. นนทบุรี' ครับ"}

            matched_df, err = search_excel_all_sheets(DIRECTOR_FILE, search_term)
            if err:
                return {"type": "text", "text": err}

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
                return {"type": "text", "text": results_text.strip()}
            else:
                return {"type": "text", "text": "พิมพ์อะไรผิดไปรึเปล่า ลองดูใหม่ดิ วุ้วววววว!!"}

        # 5. ค้นหาข้อมูลทั่วไปจากไฟล์ data.xlsx
        matched_df, err = search_excel_all_sheets(EXCEL_FILE, cleaned_text)
        if err:
            return {"type": "text", "text": err}

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
            return {"type": "text", "text": response_text.strip()}

        return {"type": "text", "text": "พิมพ์อะไรผิดไปรึเปล่า ลองดูใหม่ดิ วุ้วววววว!!"}

    except Exception as e:
        return {"type": "text", "text": f"⚠️ เกิดข้อผิดพลาด: {str(e)}"}

# Route สำหรับเสิร์ฟไฟล์รูปภาพสลิปให้ LINE ดึงไปแสดงผล
@app.route("/image/<filename>", methods=['GET'])
def send_image(filename):
    return send_from_directory('.', filename)

# Route สำหรับเช็กสถานะบอท
@app.route("/", methods=['GET'])
def health_check():
    return "LINE Bot is running fine!", 200

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

# ทำงานเมื่อมีข้อความเข้าใน LINE
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_text = event.message.text.strip()
    base_url = request.host_url  # รองรับ URL อัตโนมัติ (เช่น ผ่าน ngrok)
    result = search_data(user_text, base_url)

    if result["type"] == "image":
        line_bot_api.reply_message(
            event.reply_token,
            [
                TextSendMessage(text=result["text"]),
                ImageSendMessage(
                    original_content_url=result["original_url"],
                    preview_image_url=result["preview_url"]
                )
            ]
        )
    else:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=result["text"])
        )

if __name__ == "__main__":
    app.run(port=5000, debug=True)