import os
import logging
import asyncio
import json
import tempfile
import re
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton, BotCommand
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    ContextTypes, ConversationHandler, filters, CallbackQueryHandler
)
from docx import Document
import nest_asyncio

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Состояния
SELECT_TEMPLATE, ASKING, CONFIRMING, BLOCK_INPUT, BLOCK_CONFIRM = range(5)
PROFILE_PATH = "user_profile.json"

# Вопросы для инспекции
questions = [
    "Выберите или введите наименование товара",
    "Введите массу партии в тоннах",
    "Введите количество мест",
    "Введите транспортное средство",
    "Введите дату и номер контракта/распоряжения",
    "Введите наименование отправителя",
    "Введите сопроводительные документы (инвойс и CMR)",
    "Введите дополнительные сведения",
    "Введите дату исходящего письма и инспекции"
]

# Сопоставление с переменными в Word
mapping_keys = [
    "{{TNVED_CODE}}", "{{WEIGHT}}", "{{PLACES}}", "{{VEHICLE}}", "{{CONTRACT_INFO}}",
    "{{SENDER}}", "{{DOCS}}", "{{EXTRA_INFO}}", "{{DATE}}", "{{PRODUCT_NAME}}"
]

# Справочник ТН ВЭД
product_to_tnved = {
    "лук": "0703101900", "помидор": "0702000000", "томат": "0702000000",
    "огурец": "0707009000", "перец": "0709601000", "морковь": "0706101000",
    "капуста": "0701909000", "яблоко": "0808108000", "груша": "0808209000",
    "инжир": "0804200000", "арбуз": "0807110000", "виноград": "0806101000"
}
def detect_tnved_code(name):
    name = name.lower()
    for key, code in product_to_tnved.items():
        if key in name:
            return code
    return "0808108000"

def reorder_answers(raw):
    return [
        raw[0], raw[2], raw[3], raw[4], raw[5],
        raw[6], raw[7], raw[8], raw[9], raw[1],
    ]

def save_profile(data):
    with open(PROFILE_PATH, "w", encoding="utf-8") as f:
        json.dump(dict(zip(mapping_keys, data)), f, ensure_ascii=False, indent=2)

def replace_all(doc, replacements):
    def replace_in_paragraph(p):
        full_text = "".join(run.text for run in p.runs)
        for k, v in replacements.items():
            if k in full_text:
                full_text = full_text.replace(k, v)
        if p.runs:
            p.runs[0].text = full_text
            for i in range(1, len(p.runs)):
                p.runs[i].text = ""

    for p in doc.paragraphs:
        replace_in_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)

def generate_inspection_doc(data):
    doc = Document("Заявка на проведение инспекции.docx")
    replace_all(doc, dict(zip(mapping_keys, data)))
    out = tempfile.mktemp(suffix=".docx")
    doc.save(out)
    return out

def generate_statement_doc(blocks):
    doc = Document("Заявление на осмотр.docx")
    replace_all(doc, {"{{BLOCKS}}": "\n".join(blocks)})
    out = tempfile.mktemp(suffix=".docx")
    doc.save(out)
    return out
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    reply_markup = ReplyKeyboardMarkup([
        ["📦 Заявка на проведение инспекции", "📄 Заявление на осмотр"]
    ], resize_keyboard=True)
    await update.message.reply_text("Выберите шаблон:", reply_markup=reply_markup)
    return SELECT_TEMPLATE

async def select_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.lower()
    if "инспекц" in text:
        context.user_data["template"] = "inspection"
        context.user_data["answers"] = []
        context.user_data["step"] = 0

        # Если есть кэш
        if os.path.exists(PROFILE_PATH):
            with open(PROFILE_PATH, "r", encoding="utf-8") as f:
                context.user_data["cached"] = json.load(f)
            reply_markup = ReplyKeyboardMarkup([["✅ Да", "✏ Ввести заново"]], resize_keyboard=True)
            await update.message.reply_text("🧠 Использовать данные из последней заявки?", reply_markup=reply_markup)
            return CONFIRMING
        else:
            return await prompt_product_choice(update, context)
    else:
        context.user_data["template"] = "statement"
        context.user_data["blocks"] = []
        context.user_data["block_step"] = 0
        await update.message.reply_text("Введите госномер:")
        return BLOCK_INPUT

async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.lower()

    # Используем кэш ТОЛЬКО если пользователь находится в самом начале (step == 0)
    if "да" in text and context.user_data.get("step") == 0 and "cached" in context.user_data:
        answers = list(context.user_data["cached"].values())
    else:
        answers = context.user_data.get("answers", [])

    reordered = reorder_answers(answers)
    save_profile(reordered)
    file = generate_inspection_doc(reordered)
    await update.message.reply_document(document=open(file, "rb"))
    return ConversationHandler.END

async def prompt_product_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[InlineKeyboardButton(name.capitalize(), callback_data=name)]
                for name in list(product_to_tnved.keys())[:6]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Выберите товар или введите вручную:", reply_markup=reply_markup)
    return ASKING

async def handle_inline_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    return await process_step(query.message, context, query.data)
async def ask_question(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await process_step(update.message, context, update.message.text)

async def process_step(msg, context, text):
    step = context.user_data["step"]
    answers = context.user_data["answers"]

    if step == 0:
        tnved_code = detect_tnved_code(text.strip())
        answers.append(tnved_code)
        answers.append(text.strip())
    else:
        answers.append(text.strip())

    context.user_data["step"] += 1

    if context.user_data["step"] < len(questions):
        await msg.reply_text(questions[context.user_data["step"]])
        return ASKING
    else:
        reordered = reorder_answers(answers)
        save_profile(reordered)
        file = generate_inspection_doc(reordered)
        await msg.reply_document(document=open(file, "rb"))
        return ConversationHandler.END

# === ЛОГИКА ДЛЯ ЗАЯВЛЕНИЯ НА ОСМОТР ===

async def block_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    step = context.user_data.get("block_step", 0)
    if step == 0:
        context.user_data["v"] = update.message.text.strip()
        context.user_data["block_step"] = 1
        await update.message.reply_text("Введите документы:")
        return BLOCK_INPUT
    elif step == 1:
        context.user_data["d"] = update.message.text.strip()
        context.user_data["block_step"] = 2
        await update.message.reply_text("Введите товар:")
        return BLOCK_INPUT
    else:
        product = update.message.text.strip()
        context.user_data["blocks"].append(f"г/н {context.user_data['v']} по {context.user_data['d']}, товар: {product}")
        context.user_data["block_step"] = 0
        reply_markup = ReplyKeyboardMarkup([["➕ Да", "✅ Нет"]], resize_keyboard=True)
        await update.message.reply_text("Добавить ещё?", reply_markup=reply_markup)
        return BLOCK_CONFIRM

async def confirm_blocks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "да" in update.message.text.lower():
        await update.message.reply_text("Введите госномер:")
        return BLOCK_INPUT
    else:
        file = generate_statement_doc(context.user_data["blocks"])
        await update.message.reply_document(document=open(file, "rb"))
        return ConversationHandler.END

async def run():
    token = os.getenv("BOT_TOKEN")
    app = ApplicationBuilder().token(token).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            SELECT_TEMPLATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, select_template)],
            CONFIRMING: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm)],
            ASKING: [
                CallbackQueryHandler(handle_inline_selection),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ask_question)
            ],
            BLOCK_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, block_input)],
            BLOCK_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_blocks)],
        },
        fallbacks=[CommandHandler("start", start)],
    )

    app.add_handler(conv)
    await app.bot.set_my_commands([BotCommand("start", "Начать заполнение заявки")])
    await app.run_polling()

if __name__ == '__main__':
    nest_asyncio.apply()
    asyncio.run(run())
