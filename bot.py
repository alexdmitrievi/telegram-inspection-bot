import os
import asyncio
import nest_asyncio
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print("▶️ Получена команда /start")
    await update.message.reply_text("Бот работает ✅")

async def run():
    app = ApplicationBuilder().token(os.environ["BOT_TOKEN"]).build()
    await app.bot.delete_webhook(drop_pending_updates=True)
    app.add_handler(CommandHandler("start", start))
    await app.run_polling()

if __name__ == '__main__':
    nest_asyncio.apply()
    asyncio.run(run())

