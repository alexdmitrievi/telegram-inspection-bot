import os
import logging
import asyncio
import json
import tempfile
import re
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton, BotCommand
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    ContextTypes, ConversationHandler, filters
)
from docx import Document
from docx.shared import RGBColor
import nest_asyncio

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Справочник товаров → ТН ВЭД
product_to_tnved = {
    "лук": "0703101900",
    "помидор": "0702000000",
    "томат": "0702000000",
    "капуста": "0701909000",
    "капуста белокочанная": "0704901000",
    "огурец": "0707009000",
    "редис": "0706109000",
    "морковь": "0706101000",
    "перец": "0709601000",
    "картофель": "0701905000",
    "баклажан": "0709300000",
    "свекла": "0706109000",
    "кукуруза": "0709909000",
    "кабачок": "0709909000",
    "патиссон": "0709909000",
    "фасоль": "0708200000",
    "чеснок": "0703200000",
    "зелень": "0709990000",
    "шпинат": "0709700000",
    "кинза": "0709990000",
    "укроп": "0709990000",
    "виноград": "0806101000",
    "черешня": "0809201000",
    "вишня": "0809290000",
    "дыня": "0807190000",
    "арбуз": "0807110000",
    "яблоко": "0808108000",
    "груша": "0808209000",
    "айва": "0808400000",
    "слива": "0809400000",
    "абрикос": "0809100000",
    "персик": "0809300000",
    "инжир": "0804200000",
    "хурма": "0810907500",
    "лимон": "0805500000",
    "мандарины": "0805201000"
}

questions = [
    "Введите наименование товара",                      # step 0
    "Введите массу партии в тоннах",                    # step 1
    "Введите количество мест",                          # step 2
    "Введите транспортное средство",                    # step 3
    "Введите дату и номер контракта/распоряжения",      # step 4
    "Введите наименование отправителя",                 # step 5
    "Введите сопроводительные документы (инвойс и CMR)",# step 6
    "Введите дополнительные сведения",                  # step 7
    "Введите дату исходящего письма и инспекции"        # step 8
]

mapping_keys = [
    "{{TNVED_CODE}}", "{{WEIGHT}}", "{{PLACES}}", "{{VEHICLE}}", "{{CONTRACT_INFO}}",
    "{{SENDER}}", "{{DOCS}}", "{{EXTRA_INFO}}", "{{DATE}}", "{{PRODUCT_NAME}}"
]

profile_path = "user_profile.json"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['answers'] = []
    context.user_data['step'] = 0
    await update.message.reply_text(questions[0], reply_markup=ReplyKeyboardMarkup(
        [[KeyboardButton("🔄 Перезапустить бота")]], resize_keyboard=True))
    return ASKING

async def ask_question(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "🔄 Перезапустить бота":
        return await start(update, context)

    step = context.user_data['step']
    answers = context.user_data['answers']

    if step == 0:
        product_name = text.strip()
        tnved_code = detect_tnved_code(product_name)
        answers.append(tnved_code)        # {{TNVED_CODE}}
        answers.append(product_name)      # {{PRODUCT_NAME}}
    else:
        answers.append(validate_input(text, step))

    context.user_data['step'] += 1

    if context.user_data['step'] < len(questions):
        await update.message.reply_text(questions[context.user_data['step']])
        return ASKING
    else:
        summary = "\n".join([
            f"{questions[i]}\n➡ {answers[i+1 if i == 0 else i]}"
            for i in range(len(questions))
        ])
        await update.message.reply_text(f"Проверьте введённые данные:\n\n{summary}\n\nОтправить документы? (да/нет)")
        return CONFIRMING

def detect_tnved_code(name):
    name = name.lower()
    for keyword, code in product_to_tnved.items():
        if keyword in name:
            return code
    return "0808108000"  # fallback — яблоко

async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.lower()
    if "да" in text:
        answers = context.user_data['answers']
        save_profile(answers)
        output_files = generate_docs(answers)
        for path in output_files:
            await update.message.reply_document(document=open(path, 'rb'))
        return ConversationHandler.END
    else:
        await update.message.reply_text("Ок, начнём заново.")
        context.user_data['answers'] = []
        context.user_data['step'] = 0
        return ASKING

def validate_input(text, step):
    try:
        if step == 1:
            return re.sub(r"[^0-9.,]", "", text).replace(",", ".")
        elif step == 2:
            return re.sub(r"\D", "", text)
        elif step == 8:
            d = re.search(r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}", text)
            return datetime.strptime(d.group(), "%d.%m.%Y").strftime("%d.%m.%Y") if d else text
        elif step == 5:
            return text.upper()
        else:
            return text.strip()
    except Exception as e:
        logger.error(f"Ошибка валидации: {e}")
        return text.strip()

def save_profile(answers):
    try:
        with open("user_profile.json", "w", encoding="utf-8") as f:
            json.dump(dict(zip(mapping_keys, answers)), f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"Ошибка при сохранении профиля: {e}")

def generate_docs(answers):
    replacements = dict(zip(mapping_keys, answers))
    result_files = []
    for template_path in ["Заявка на проведение инспекции.docx", "Заявление на осмотр.docx"]:
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
        output_path = tempfile.mktemp(suffix=".docx")
        doc.save(output_path)
        result_files.append(output_path)
    return result_files

async def run():
    TOKEN = os.getenv("BOT_TOKEN")
    app = ApplicationBuilder().token(TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ASKING: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_question)],
            CONFIRMING: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm)],
        },
        fallbacks=[MessageHandler(filters.Regex("🔄 Перезапустить бота"), start)],
    )

    app.add_handler(conv)
    app.add_handler(MessageHandler(filters.ALL, log_all_updates))

    await app.bot.set_my_commands([BotCommand("start", "Начать заполнение заявки")])
    await app.run_polling()

if __name__ == '__main__':
    try:
        nest_asyncio.apply()
        asyncio.get_event_loop().run_until_complete(run())
    except Exception as e:
        logger.error(f"Критическая ошибка запуска: {e}")
