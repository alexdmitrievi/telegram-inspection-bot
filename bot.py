import os
import logging
import asyncio
import json
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    ContextTypes, ConversationHandler, filters, CallbackQueryHandler
)
from docx import Document
import nest_asyncio

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

SELECT_TEMPLATE, ASKING, CONFIRMING, BLOCK_INPUT, BLOCK_CONFIRM = range(5)
PROFILE_PATH = "user_profile.json"

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

mapping_keys = [
    "{{TNVED_CODE}}", "{{WEIGHT}}", "{{PLACES}}", "{{VEHICLE}}", "{{CONTRACT_INFO}}",
    "{{SENDER}}", "{{DOCS}}", "{{EXTRA_INFO}}", "{{DATE}}", "{{PRODUCT_NAME}}"
]

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

def save_profile(data):
    with open(PROFILE_PATH, "w", encoding="utf-8") as f:
        json.dump(dict(zip(mapping_keys, data)), f, ensure_ascii=False, indent=2)

def replace_all(doc, replacements):
    for p in doc.paragraphs:
        for k, v in replacements.items():
            if k in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if k in inline[i].text:
                        inline[i].text = inline[i].text.replace(k, v)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_all(cell, replacements)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    reply_markup = ReplyKeyboardMarkup([
        ["\U0001F4E6 Заявка на проведение инспекции", "\U0001F4C4 Заявление на осмотр"]
    ], resize_keyboard=True)
    await update.message.reply_text("Выберите шаблон:", reply_markup=reply_markup)
    return SELECT_TEMPLATE

async def select_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.lower()

    if "инспекц" in text:
        context.user_data.clear()
        context.user_data["template"] = "inspection"
        context.user_data["answers"] = []
        context.user_data["step"] = 0

        if os.path.exists(PROFILE_PATH):
            with open(PROFILE_PATH, "r", encoding="utf-8") as f:
                context.user_data["cached"] = json.load(f)
            reply_markup = ReplyKeyboardMarkup([["\u2705 Да", "\u270F Ввести заново"]], resize_keyboard=True)
            await update.message.reply_text("\U0001F9E0 Использовать данные из последней заявки?", reply_markup=reply_markup)
            return CONFIRMING
        else:
            return await prompt_product_choice(update, context)

    elif "ввести заново" in text:
        context.user_data["answers"] = []
        context.user_data["step"] = 0
        return await prompt_product_choice(update, context)

    else:
        context.user_data["template"] = "statement"
        context.user_data["blocks"] = []
        context.user_data["block_step"] = 0
        await update.message.reply_text("Введите госномер:")
        return BLOCK_INPUT

async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().lower()
    answers = context.user_data.get("answers", {})

    if text in ["да", "\u2705 да"] and context.user_data.get("step") == 0 and "cached" in context.user_data:
        answers = context.user_data["cached"]
        context.user_data["answers"] = answers

    save_profile(answers)
    file = generate_inspection_doc_from_dict(answers)
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
    step = context.user_data.get("step", 0)
    if not isinstance(context.user_data.get("answers"), dict):
        context.user_data["answers"] = {}
    current_answers = context.user_data["answers"]

    key_order = [
        "{{PRODUCT_NAME}}", "{{WEIGHT}}", "{{PLACES}}", "{{VEHICLE}}",
        "{{CONTRACT_INFO}}", "{{SENDER}}", "{{DOCS}}", "{{EXTRA_INFO}}", "{{DATE}}"
    ]

    if step == 0:
        product = text.strip()
        tnved = detect_tnved_code(product)
        current_answers["{{PRODUCT_NAME}}"] = product
        current_answers["{{TNVED_CODE}}"] = tnved
    else:
        key = key_order[step]
        current_answers[key] = text.strip()

    context.user_data["answers"] = current_answers
    context.user_data["step"] = step + 1

    if context.user_data["step"] < len(questions):
        await msg.reply_text(questions[context.user_data["step"]])
        return ASKING
    else:
        summary = "\n".join([
            f"{questions[i]}: {current_answers.get(key_order[i], '—')}"
            for i in range(len(questions))
        ])
        await msg.reply_text(
            f"Проверьте введённые данные:\n\n{summary}\n\nОтправить документы? (да/нет)",
            reply_markup=ReplyKeyboardMarkup([["\U0001F501 Перезапустить", "да", "нет"]], resize_keyboard=True)
        )
        return CONFIRMING

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

    elif step == 2:
        context.user_data["product"] = update.message.text.strip()
        if "statement_date" not in context.user_data:
            context.user_data["ask_date"] = True
            await update.message.reply_text("Введите дату заявления и осмотра:")
            return BLOCK_INPUT
        else:
            context.user_data["blocks"].append({
                "{{VEHICLE}}": context.user_data["v"],
                "{{DOCS}}": context.user_data["d"],
                "{{PRODUCT_NAME}}": context.user_data["product"]
            })
            context.user_data["block_step"] = 0
            reply_markup = ReplyKeyboardMarkup([["\u2795 Да", "\u2705 Нет"]], resize_keyboard=True)
            await update.message.reply_text("Добавить ещё?", reply_markup=reply_markup)
            return BLOCK_CONFIRM

    elif context.user_data.get("ask_date"):
        context.user_data["statement_date"] = update.message.text.strip()
        context.user_data["blocks"].append({
            "{{VEHICLE}}": context.user_data["v"],
            "{{DOCS}}": context.user_data["d"],
            "{{PRODUCT_NAME}}": context.user_data["product"]
        })
        context.user_data["block_step"] = 0
        context.user_data.pop("ask_date", None)
        reply_markup = ReplyKeyboardMarkup([["\u2795 Да", "\u2705 Нет"]], resize_keyboard=True)
        await update.message.reply_text("Добавить ещё?", reply_markup=reply_markup)
        return BLOCK_CONFIRM

async def confirm_blocks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "да" in update.message.text.lower():
        await update.message.reply_text("Введите госномер:")
        return BLOCK_INPUT
    else:
        file = generate_statement_doc(context.user_data["blocks"], context.user_data.get("statement_date", "—"))
        await update.message.reply_document(document=open(file, "rb"))
        return ConversationHandler.END

def generate_statement_doc(blocks, date):
    template_path = "Заявление на осмотр.docx"
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"Заявление_на_осмотр_{timestamp}.docx")
    doc = Document(template_path)

    for block in blocks:
        replace_all(doc, block)

    replace_all(doc, {"{{DATE}}": date})
    doc.save(output_path)
    return output_path

def generate_inspection_doc_from_dict(replacements):
    template_path = "Заявка на проведение инспекции.docx"
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"Заявка_на_проведение_инспекции_{timestamp}.docx")
    doc = Document(template_path)
    replace_all(doc, replacements)
    doc.save(output_path)
    return output_path

async def run():
    app = ApplicationBuilder().token("7548023133:AAFfDrnLlF340dAfqrhfjfs8UF4_4NG7f84").build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            SELECT_TEMPLATE: [MessageHandler(filters.TEXT, select_template)],
            CONFIRMING: [MessageHandler(filters.TEXT, confirm)],
            ASKING: [
                MessageHandler(filters.TEXT & (~filters.COMMAND), ask_question),
                CallbackQueryHandler(handle_inline_selection)
            ],
            BLOCK_INPUT: [MessageHandler(filters.TEXT, block_input)],
            BLOCK_CONFIRM: [MessageHandler(filters.TEXT, confirm_blocks)],
        },
        fallbacks=[CommandHandler("start", start)],
    )

    app.add_handler(conv_handler)
    await app.run_polling()

if __name__ == '__main__':
    nest_asyncio.apply()
    asyncio.run(run())

