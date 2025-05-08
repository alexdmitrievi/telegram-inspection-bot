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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

(ASKING, CONFIRMING) = range(2)

# Вопросы для анкеты
questions = [
    "Введите код ТН ВЭД",
    "Введите массу партии в тоннах",
    "Введите количество мест",
    "Введите дату и номер контракта/распоряжения на поставку",
    "Введите наименование отправителя",
    "Введите сопроводительные документы (инвойс и CMR)",
    "Введите дополнительные сведения",
    "Введите предполагаемую дату начала инспекции",
    "Введите наименование товара",
    "Введите дату исходящего письма"
]

# Переменные в шаблонах
mapping_keys = [
    "{{TNVED_CODE}}", "{{WEIGHT}}", "{{PLACES}}", "{{CONTRACT_INFO}}",
    "{{SENDER}}", "{{DOCS}}", "{{EXTRA_INFO}}", "{{INSPECTION_DATE}}",
    "{{PRODUCT_NAME}}", "{{DATE}}"
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
    answer = validate_input(text, step)
    context.user_data['answers'].append(answer)
    context.user_data['step'] += 1

    if context.user_data['step'] < len(questions):
        await update.message.reply_text(questions[context.user_data['step']])
        return ASKING
    else:
        summary = "\n".join([
            f"{questions[i]}\n➡ {context.user_data['answers'][i]}"
            for i in range(len(questions))
        ])
        await update.message.reply_text(f"Проверьте введённые данные:\n\n{summary}\n\nОтправить документы? (да/нет)")
        return CONFIRMING

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
        await update.message.reply_text("Ок, начнём заново. Введите первую информацию:")
        context.user_data['answers'] = []
        context.user_data['step'] = 0
        return ASKING

def validate_input(text, step):
    try:
        if step == 1:  # масса
            return re.sub(r"[^0-9.,]", "", text).replace(",", ".")
        elif step == 2:  # кол-во мест
            return re.sub(r"\D", "", text)
        elif step in [7, 9]:  # даты
            d = re.search(r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}", text)
            return datetime.strptime(d.group(), "%d.%m.%Y").strftime("%d.%m.%Y") if d else text
        elif step == 4:  # отправитель
            return text.upper()
        else:
            return text.strip().capitalize()
    except Exception as e:
        logger.error(f"Ошибка валидации: {e}")
        return text.strip()

def save_profile(answers):
    try:
        with open(profile_path, 'w', encoding='utf-8') as f:
            json.dump({k: v for k, v in zip(mapping_keys, answers)}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"Ошибка при сохранении профиля: {e}")

def generate_docs(answers):
    replacements = dict(zip(mapping_keys, answers))
    template_files = ["Заявка на проведение инспекции.docx", "Заявление на осмотр.docx"]
    result_files = []

    for template_path in template_files:
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

    await app.bot.set_my_commands([BotCommand("start", "Начать заполнение заявки")])

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ASKING: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_question)],
            CONFIRMING: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm)]
        },
        fallbacks=[MessageHandler(filters.Regex("🔄 Перезапустить бота"), start)]
    )

    app.add_handler(conv)

    await app.initialize()
    await app.start()
    await asyncio.Event().wait()

if __name__ == '__main__':
    try:
        asyncio.run(run())
    except Exception as e:
        logger.error(f"Критическая ошибка запуска: {e}")