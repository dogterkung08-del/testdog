import os
import re
import pandas as pd
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    filters,
)

# กำหนดชื่อไฟล์ Excel
EXCEL_FILE = "data.xlsx"
DIRECTOR_FILE = "name.xlsx"  # ไฟล์เก็บข้อมูล ผอ. แต่ละจังหวัด


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
  user_text = update.message.text.strip()

  try:
    # 1. ค้นหาข้อมูล ผอ. จากไฟล์ name.xlsx
    if "ผอ" in user_text:
      # ใช้ Regular Expression ตัดคำว่า "ผอ.", "ผอ", "ของ" ออก
      # และลบช่องว่างส่วนเกิน ไม่ว่าจะพิมพ์ติดกันหรือเว้นวรรคกว้างแค่ไหน
      search_term = re.sub(r"ผอ\.|ผอ|ของ", "", user_text).strip()

      if not search_term:
        search_term = user_text

      # ตรวจสอบว่ามีไฟล์ name.xlsx หรือไม่
      if not os.path.exists(DIRECTOR_FILE):
        await update.message.reply_text(
            f"⚠️ ไม่พบไฟล์ {DIRECTOR_FILE} ในระบบ กรุณาตรวจสอบชื่อไฟล์ครับ"
        )
        return

      # อ่านไฟล์ name.xlsx
      df_dir = pd.read_excel(DIRECTOR_FILE, dtype=str).fillna("")

      # ค้นหาในคอลัมน์ A (ชื่อสถานพินิจฯ) หรือ คอลัมน์ B (ชื่อ ผอ.)
      # กำหนด regex=False เพื่อป้องกัน Error เมื่อพิมพ์สัญลักษณ์พิเศษ
      mask = df_dir.iloc[:, 0].str.contains(
          search_term, regex=False, case=False, na=False
      ) | df_dir.iloc[:, 1].str.contains(
          search_term, regex=False, case=False, na=False
      )
      matched_rows = df_dir[mask]

      if not matched_rows.empty:
        results_text = f"📋 **ผลการค้นหาข้อมูล ผอ. ({search_term}):**\n\n"

        for _, row in matched_rows.iterrows():
          org_name = row.iloc[0] if len(row) > 0 else ""
          director_name = row.iloc[1] if len(row) > 1 else ""
          updated_date = row.iloc[2] if len(row) > 2 else ""

          results_text += f"🏢 **{org_name}**\n"
          results_text += f"👤 **ผอ.:** {director_name}\n"
          if updated_date:
            results_text += f"📅 **อัปเดตเมื่อ:** {updated_date}\n"
          results_text += "------------------------------\n"

        await update.message.reply_text(results_text.strip())
      else:
        # เปลี่ยนข้อความเมื่อหาไม่พบ
        await update.message.reply_text(
            "พิมพ์อะไรผิดไปรึเปล่า ลองดูใหม่ดิ วุ้วววววว!!"
        )
      return

    # 2. ค้นหาข้อมูลทั่วไปจากไฟล์ data.xlsx
    if not os.path.exists(EXCEL_FILE):
      await update.message.reply_text(
          f"⚠️ ไม่พบไฟล์ {EXCEL_FILE} ในระบบ กรุณาตรวจสอบชื่อไฟล์ครับ"
      )
      return

    df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")

    # ใส่ regex=False ป้องกัน Error เวลาผู้ใช้พิมพ์ ??? หรือสัญลักษณ์อื่น ๆ
    mask = (
        df.astype(str)
        .apply(
            lambda x: x.str.contains(
                user_text, regex=False, case=False, na=False
            )
        )
        .any(axis=1)
    )
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
        response_text += (
            f"\n*(แสดง 10 รายการแรก จากทั้งหมด {len(matched_rows)} รายการ)*"
        )

      await update.message.reply_text(response_text.strip())
    else:
      # เปลี่ยนข้อความเมื่อหาไม่พบ
      await update.message.reply_text(
          "พิมพ์อะไรผิดไปรึเปล่า ลองดูใหม่ดิ วุ้วววววว!!"
      )

  except Exception:
    # เปลี่ยนข้อความแจ้งเตือน Error ทั้งหมดเป็นภาษาตามต้องการ
    await update.message.reply_text("อิหยัง ？")


if __name__ == "__main__":
  TOKEN = "8909263779:AAG2gEucHIZp4zDglhpJw5MP2MNSDb7EOlw"

  app = ApplicationBuilder().token(TOKEN).build()

  app.add_handler(
      MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message)
  )

  print("🤖 Telegram Bot กำลังทำงาน...")
  app.run_polling()