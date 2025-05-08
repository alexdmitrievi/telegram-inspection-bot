import os
import logging
import tempfile
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton, BotCommand
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    ContextTypes, filters, ConversationHandler
)
from docx import Document
from docx.shared import RGBColor

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TEMPLATES = [
    "Заявка_на_проведение_инспекции.docx",
    "Заявление_на_осмотр.docx"
]

FILLING = 0

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["template_index"] = 0
    context.user_data["replacements"] = {}
    context.user_data["fields"] = []
    context.user_data["current"] = 0

    await update.message.reply_text("Добро пожаловать! Сейчас вы будете поочерёдно заполнять два шаблона.")
    return await process_next_template(update, context)

async def process_next_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if context.user_data["template_index"] >= len(TEMPLATES):
        await update.message.reply_text("Все документы заполнены.")
        return ConversationHandler.END

    template_path = TEMPLATES[context.user_data["template_index"]]
    fields = extract_red_text(template_path)
    context.user_data["fields"] = fields
    context.user_data["current"] = 0
    context.user_data["template_path"] = template_path
    context.user_data["replacements"] = {}

    if not fields:
        await update.message.reply_text(f"В шаблоне «{template_path}» не найдено полей для заполнения.")
        context.user_data["template_index"] += 1
        return await process_next_template(update, context)

    await update.message.reply_text(f"Шаблон: {template_path}. Найдено полей: {len(fields)}.")
    return await ask_next_field(update, context)

async def ask_next_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    fields = context.user_data["fields"]
    current = context.user_data["current"]

    if current >= len(fields):
        output_path = fill_docx(context.user_data["template_path"], context.user_data["replacements"])
        await update.message.reply_document(document=open(output_path, "rb"))
        context.user_data["template_index"] += 1
        return await process_next_template(update, context)

    field = fields[current]
    await update.message.reply_text(f"Введите значение для поля: «{field}»")
    return FILLING

async def receive_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    current = context.user_data["current"]
    fields = context.user_data["fields"]
    value = update.message.text
    context.user_data["replacements"][fields[current]] = value
    context.user_data["current"] += 1
    return await ask_next_field(update, context)

def extract_red_text(path):
    doc = Document(path)
    fields = set()

    def collect_red_runs(paragraphs):
        for p in paragraphs:
            for run in p.runs:
                if run.font.color and run.font.color.rgb == RGBColor(255, 0, 0):
                    fields.add(run.text.strip())

    for para in doc.paragraphs:
        collect_red_runs([para])

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                collect_red_runs(cell.paragraphs)

    return list(fields)

def fill_docx(template_path, replacements):
    doc = Document(template_path)

    def replace_runs(paragraphs):
        for p in paragraphs:
            for run in p.runs:
                if run.font.color and run.font.color.rgb == RGBColor(255, 0, 0):
                    for key, val in replacements.items():
                        if run.text.strip() == key:
                            run.text = val

    for para in doc.paragraphs:
        replace_runs([para])

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_runs(cell.paragraphs)

    output_path = tempfile.mktemp(suffix=".docx")
    doc.save(output_path)
    return output_path

def main():
    TOKEN = os.getenv("BOT_TOKEN")
    app = ApplicationBuilder().token(TOKEN).build()

    app.bot.set_my_commands([
        BotCommand("start", "Начать заполнение шаблонов")
    ])

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            FILLING: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_field)],
        },
        fallbacks=[],
    )

    app.add_handler(conv)
    app.run_polling()

if __name__ == "__main__":
    main()