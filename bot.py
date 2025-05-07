# Финальный Telegram-бот с поддержкой PDF, Excel и генерацией двух шаблонов

import os
import logging
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters, ConversationHandler
from docx import Document
from docx.shared import RGBColor
import pytesseract
from PIL import Image
import tempfile
import pdfplumber
import openpyxl

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

UPLOAD, PROCESS = range(2)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[KeyboardButton("🔄 Перезапустить бота")]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Добро пожаловать! Пожалуйста, отправьте инвойс, CMR или TIR.", reply_markup=reply_markup)
    return UPLOAD

async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await start(update, context)

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = update.message.document or update.message.photo[-1]
    file_path = tempfile.mktemp()
    new_file = await file.get_file()
    await new_file.download_to_drive(file_path)

    text = extract_text(file_path)
    logger.info("Извлечённый текст:\n%s", text)

    replacements = {
        'ЛУК РЕПЧАТЫЙ СВЕЖИЙ, УРОЖАЙ 2025 г.': find_line_containing(text, 'лук') or 'Лук репчатый свежий, урожай 2025 г.',
        '0703101900': find_code(text),
        '23,220': find_mass(text),
        '01W353JC/017827BA': find_vehicle_number(text),
        'ROM-2 от 23.04.2025 г.': find_contract(text),
        'ООО «ROMA TRADE»': find_sender(text),
        'ИНВОЙС RTRZ-64 от 03.05.2025': find_invoice(text),
    }

    out1 = fill_docx_by_color("Заявка на проведение инспекции лук 353.docx", replacements)
    out2 = fill_docx_by_color("Заявление на осмотр 354 153.docx", replacements)

    await update.message.reply_document(document=open(out1, 'rb'), filename="Заявка_на_проведение_инспекции.docx")
    await update.message.reply_document(document=open(out2, 'rb'), filename="Заявление_на_осмотр.docx")

    return PROCESS

def extract_text(file_path):
    ext = os.path.splitext(file_path)[-1].lower()
    if ext in ['.jpg', '.jpeg', '.png']:
        return pytesseract.image_to_string(Image.open(file_path), lang='rus+eng')
    elif ext.endswith('.pdf'):
        text = ""
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        return text
    elif ext.endswith('.xlsx'):
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        values = []
        for row in sheet.iter_rows(values_only=True):
            values.extend([str(cell) for cell in row if cell])
        return " ".join(values)
    return ""

def find_line_containing(text, keyword):
    for line in text.splitlines():
        if keyword.lower() in line.lower():
            return line.strip()
    return None

def find_code(text):
    import re
    match = re.search(r'07\d{6,}', text)
    return match.group(0) if match else '0703101900'

def find_mass(text):
    import re
    match = re.search(r'\b(2\d{4,5})\b', text)
    return match.group(1) if match else '23220'

def find_vehicle_number(text):
    match = find_line_containing(text, 'W')
    return match if match else '01W353JC/017827BA'

def find_contract(text):
    return find_line_containing(text, 'контракт') or 'ROM-2 от 23.04.2025 г.'

def find_sender(text):
    return find_line_containing(text, 'ROMA TRADE') or 'ООО «ROMA TRADE»'

def find_invoice(text):
    return find_line_containing(text, 'инвойс') or 'ИНВОЙС RTRZ-64 от 03.05.2025'

def fill_docx_by_color(template_path, replacements):
    doc = Document(template_path)
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.color and run.font.color.rgb == RGBColor(255, 0, 0):
                for key, val in replacements.items():
                    if key in run.text:
                        run.text = run.text.replace(key, val)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.font.color and run.font.color.rgb == RGBColor(255, 0, 0):
                            for key, val in replacements.items():
                                if key in run.text:
                                    run.text = run.text.replace(key, val)
    output_path = tempfile.mktemp(suffix='.docx')
    doc.save(output_path)
    return output_path

def main():
    import asyncio
    import os

    TOKEN = os.getenv("BOT_TOKEN")
    WEBHOOK_URL = os.getenv("WEBHOOK_URL")

    app = ApplicationBuilder().token(TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            UPLOAD: [
                MessageHandler(filters.Document.ALL | filters.PHOTO, handle_file),
                MessageHandler(filters.Regex("🔄 Перезапустить бота"), restart),
            ],
            PROCESS: [
                MessageHandler(filters.Regex("🔄 Перезапустить бота"), restart)
            ]
        },
        fallbacks=[CommandHandler("start", start)],
    )
    app.add_handler(conv)

    async def run():
        await app.initialize()
        await app.bot.set_webhook(WEBHOOK_URL)
        await app.start()
        await app.run_webhook(             # 👈 ОБЯЗАТЕЛЬНО!
            listen="0.0.0.0",
            port=10000,
            webhook_path="/webhook"
        )

    asyncio.run(run())

if __name__ == '__main__':
    main()
