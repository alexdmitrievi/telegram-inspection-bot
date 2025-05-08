
import os
import logging
import tempfile
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    ContextTypes, filters, ConversationHandler
)
from docx import Document
from docx.shared import RGBColor

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TEMPLATE_PATHS = [
    "Заявка_на_проведение_инспекции.docx",
    "Заявление_на_осмотр.docx"
]

FILLING, = range(1)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["replacements"] = {}
    context.user_data["documents"] = TEMPLATE_PATHS.copy()
    context.user_data["current_doc"] = 0
    return await process_next_document(update, context)

async def process_next_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if context.user_data["current_doc"] >= len(context.user_data["documents"]):
        await update.message.reply_text("Все шаблоны успешно заполнены!")
        return ConversationHandler.END

    current_template = context.user_data["documents"][context.user_data["current_doc"]]
    context.user_data["fields"] = extract_red_text(current_template)
    context.user_data["current_field"] = 0

    if not context.user_data["fields"]:
        await update.message.reply_text(f"В шаблоне {current_template} нет полей для заполнения.")
        context.user_data["current_doc"] += 1
        return await process_next_document(update, context)

    await update.message.reply_text(f"Заполняем шаблон: {current_template}")
    return await ask_next_field(update, context)

async def ask_next_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    fields = context.user_data["fields"]
    current = context.user_data["current_field"]

    if current >= len(fields):
        current_template = context.user_data["documents"][context.user_data["current_doc"]]
        output_path = fill_docx(current_template, context.user_data["replacements"])
        await update.message.reply_document(document=open(output_path, "rb"), filename=f"{os.path.basename(current_template)}")
        context.user_data["current_doc"] += 1
        return await process_next_document(update, context)

    field = fields[current]
    await update.message.reply_text(f"Введите значение для поля: «{field}»")
    return FILLING

async def receive_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    current = context.user_data["current_field"]
    fields = context.user_data["fields"]
    value = update.message.text
    context.user_data["replacements"][fields[current]] = value
    context.user_data["current_field"] += 1
    return await ask_next_field(update, context)

async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await start(update, context)

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

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={FILLING: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_field)]},
        fallbacks=[MessageHandler(filters.Regex("🔄 Перезапустить бота"), restart)],
    )
    app.add_handler(conv)

    async def run():
        await app.bot.set_my_commands([("start", "Запустить бота")])
        await app.initialize()
        await app.start()
        await app.updater.start_polling()
        await app.updater.idle()

    import asyncio
    asyncio.run(run())

if __name__ == "__main__":
    main()