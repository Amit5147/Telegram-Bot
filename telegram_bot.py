from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import pandas as pd
import os
import whisper
from datetime import datetime
import ffmpeg
import uuid
import asyncio
import subprocess

# Load Whisper model
model = whisper.load_model("base")

# Create empty DataFrame
if not os.path.exists("data.xlsx"):
    df = pd.DataFrame(columns=["User", "Text", "VoiceText", "Location", "Photo", "Date"])
    df.to_excel("data.xlsx", index=False)

# --- Handlers ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 Hello! Send me a message, photo, location, or voice note.")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username or update.message.from_user.first_name
    text = update.message.text
    time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    save_to_excel(user, text, "", "", "", time)
    await update.message.reply_text("✅ Text saved.")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username or update.message.from_user.first_name
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    path = f"photos/{uuid.uuid4()}.jpg"
    os.makedirs("photos", exist_ok=True)
    await file.download_to_drive(path)
    time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    save_to_excel(user, "", "", "", path, time)
    await update.message.reply_text("📸 Photo saved.")

async def handle_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username or update.message.from_user.first_name
    loc = update.message.location
    location_str = f"{loc.latitude}, {loc.longitude}"
    time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    save_to_excel(user, "", "", location_str, "", time)
    await update.message.reply_text("📍 Location saved.")

async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username or update.message.from_user.first_name
    voice = update.message.voice
    file = await context.bot.get_file(voice.file_id)
    os.makedirs("audio", exist_ok=True)
    ogg_path = os.path.abspath(f"audio/{uuid.uuid4()}.ogg")
    wav_path = ogg_path.replace(".ogg", ".wav")
    await file.download_to_drive(ogg_path)

    # Debug: Check if file exists
    if not os.path.exists(ogg_path):
        await update.message.reply_text(f"Audio file not found: {ogg_path}")
        return

    # Check if ffmpeg is available
    try:
        subprocess.run(["ffmpeg", "-version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception:
        await update.message.reply_text("FFmpeg is not installed or not in PATH. Please install FFmpeg and add it to your PATH.")
        return

    # Convert OGG to WAV
    try:
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(
            None,
            lambda: ffmpeg.input(ogg_path).output(wav_path).run(overwrite_output=True)
        )
        # Debug: Check if WAV was created
        if not os.path.exists(wav_path):
            await update.message.reply_text(f"WAV file not created: {wav_path}")
            return
    except Exception as e:
        await update.message.reply_text(f"FFmpeg error: {e}")
        return

    # Whisper transcription
    result = model.transcribe(wav_path)
    text = result["text"]
    time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    save_to_excel(user, "", text, "", "", time)
    await update.message.reply_text(f"🗣️ Voice converted: {text}")

async def send_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_document(chat_id=update.effective_chat.id, document=open("data.xlsx", "rb"))

# --- Save Data Function ---
def save_to_excel(user, text, voice_text, location, photo, time):
    df = pd.read_excel("data.xlsx")
    new_row = {"User": user, "Text": text, "VoiceText": voice_text, "Location": location, "Photo": photo, "Date": time}
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel("data.xlsx", index=False)

# --- Run Bot ---
if __name__ == "__main__":
    TOKEN = "8477518231:AAFV66k3RO6-6OCYUa3sXlUYhOk77DD59Yw"
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", send_excel))
    app.add_handler(MessageHandler(filters.TEXT, handle_text))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.LOCATION, handle_location))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))

    print("🤖 Bot is running...")
    app.run_polling()
