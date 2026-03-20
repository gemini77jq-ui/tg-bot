"""
Telegram бот для регистрации автомобилей на территорию.
Данные сохраняются в Google Таблицы.
"""

import logging
import re
from datetime import datetime
from zoneinfo import ZoneInfo
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)
from google_sheets import GoogleSheetsManager
from config import BOT_TOKEN, ADMIN_CHAT_ID

# Настройка логирования
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# Шаги диалога
CAR_BRAND, CAR_NUMBER, ARRIVAL_DATE, ARRIVAL_TIME, TEAM, CONFIRM = range(6)

# Кнопки подтверждения
CONFIRM_KEYBOARD = ReplyKeyboardMarkup(
    [["✅ Подтвердить", "❌ Отменить"]],
    resize_keyboard=True,
    one_time_keyboard=True,
)


def format_car_number(number: str) -> str:
    """Приводит номер к верхнему регистру и убирает лишние пробелы."""
    return number.strip().upper()


def is_valid_date(date_str: str) -> bool:
    """Проверяет формат даты ДД.ММ.ГГГГ."""
    try:
        datetime.strptime(date_str.strip(), "%d.%m.%Y")
        return True
    except ValueError:
        return False


def is_valid_time_range(time_str: str) -> bool:
    """Проверяет формат диапазона времени ЧЧ:ММ - ЧЧ:ММ."""
    pattern = r"^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$"
    return bool(re.match(pattern, time_str.strip()))


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Начало диалога регистрации."""
    user = update.effective_user
    logger.info(f"Пользователь {user.id} ({user.username}) начал регистрацию")

    await update.message.reply_text(
        f"👋 Приветствуем Вас!\n\n"
        "🚗 Этот бот поможет зарегистрировать ваш автомобиль для въезда на территорию.\n\n"
        "Пожалуйста, заполните данные. Это займёт около минуты.\n\n"
        "🚗 *Шаг 1 из 5* — Введите марку автомобиля:\n"
        "_Пример: Toyota, BMW, Lada_",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove(),
    )
    return CAR_BRAND


async def get_car_brand(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение марки авто."""
    brand = update.message.text.strip()

    if len(brand) < 2:
        await update.message.reply_text("⚠️ Введите корректное название марки:")
        return CAR_BRAND

    context.user_data["car_brand"] = brand.title()
    await update.message.reply_text(
        f"✅ Марка: *{brand.title()}*\n\n🔢 *Шаг 2 из 5* — Введите гос. номер автомобиля:\n"
        "_Пример: А123БВ777_",
        parse_mode="Markdown",
    )
    return CAR_NUMBER


async def get_car_number(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение гос. номера."""
    number = format_car_number(update.message.text)

    if len(number) < 4:
        await update.message.reply_text(
            "⚠️ Введите корректный гос. номер автомобиля:"
        )
        return CAR_NUMBER

    context.user_data["car_number"] = number
    await update.message.reply_text(
        f"✅ Гос. номер: `{number}`\n\n📅 *Шаг 3 из 5* — Введите дату прибытия:\n"
        "_Формат: ДД.ММ.ГГГГ, например 25.06.2025_",
        parse_mode="Markdown",
    )
    return ARRIVAL_DATE


async def get_arrival_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение даты прибытия."""
    date_str = update.message.text.strip()

    if not is_valid_date(date_str):
        await update.message.reply_text(
            "⚠️ Неверный формат даты. Введите в формате ДД.ММ.ГГГГ:\n"
            "_Пример: 25.06.2025_",
            parse_mode="Markdown",
        )
        return ARRIVAL_DATE

    context.user_data["arrival_date"] = date_str.strip()
    await update.message.reply_text(
        f"✅ Дата прибытия: *{date_str.strip()}*\n\n"
        "⏰ *Шаг 4 из 5* — Введите диапазон времени пребывания:\n"
        "_Формат: ЧЧ:ММ - ЧЧ:ММ, например 09:00 - 18:00_",
        parse_mode="Markdown",
    )
    return ARRIVAL_TIME


async def get_arrival_time(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение диапазона времени пребывания."""
    time_str = update.message.text.strip()

    if not is_valid_time_range(time_str):
        await update.message.reply_text(
            "⚠️ Неверный формат. Введите диапазон времени:\n"
            "_Формат: ЧЧ:ММ - ЧЧ:ММ, например 09:00 - 18:00_",
            parse_mode="Markdown",
        )
        return ARRIVAL_TIME

    context.user_data["arrival_time"] = time_str
    await update.message.reply_text(
        f"✅ Время пребывания: *{time_str}*\n\n"
        "🏒 *Шаг 5 из 5* — Введите название команды или школы:",
        parse_mode="Markdown",
    )
    return TEAM


async def get_team(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение названия команды/школы и показ итоговой формы."""
    team = update.message.text.strip()

    if len(team) < 2:
        await update.message.reply_text("⚠️ Введите корректное название команды или школы:")
        return TEAM

    context.user_data["team"] = team

    data = context.user_data
    summary = (
        "📋 *Проверьте введённые данные:*\n\n"
        f"🚗 Марка: {data['car_brand']}\n"
        f"🔢 Гос. номер: `{data['car_number']}`\n"
        f"📅 Дата прибытия: {data['arrival_date']}\n"
        f"⏰ Время пребывания: {data['arrival_time']}\n"
        f"🏒 Команда/Школа: {data['team']}\n\n"
        "Всё верно?"
    )

    await update.message.reply_text(
        summary,
        parse_mode="Markdown",
        reply_markup=CONFIRM_KEYBOARD,
    )
    return CONFIRM


async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Подтверждение и сохранение данных."""
    answer = update.message.text

    if answer == "❌ Отменить":
        await update.message.reply_text(
            "🚫 Регистрация отменена.\n\nДля начала заново введите /start",
            reply_markup=ReplyKeyboardRemove(),
        )
        context.user_data.clear()
        return ConversationHandler.END

    if answer != "✅ Подтвердить":
        await update.message.reply_text(
            "Пожалуйста, нажмите одну из кнопок:",
            reply_markup=CONFIRM_KEYBOARD,
        )
        return CONFIRM

    sheets = GoogleSheetsManager()
    user = update.effective_user
    data = context.user_data

    # Проверка дубликата по гос. номеру
    if sheets.is_duplicate(data["car_number"]):
        await update.message.reply_text(
            f"⚠️ Автомобиль с номером *{data['car_number']}* уже зарегистрирован!\n\n"
            "Если это ошибка, свяжитесь с администратором.",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ConversationHandler.END

    record = {
        "timestamp": datetime.now(ZoneInfo("Europe/Moscow")).strftime("%d.%m.%Y %H:%M"),
        "tg_id": str(user.id),
        "tg_username": f"@{user.username}" if user.username else "—",
        "car_brand": data["car_brand"],
        "car_number": data["car_number"],
        "arrival_date": data["arrival_date"],
        "arrival_time": data["arrival_time"],
        "team": data["team"],
        "status": "Ожидает одобрения",
    }

    success = sheets.add_record(record)

    if success:
        await update.message.reply_text(
            "✅ *Заявка успешно отправлена!*\n\n"
            f"🚗 Марка: {data['car_brand']}\n"
            f"🔢 Гос. номер: `{data['car_number']}`\n"
            f"📅 Дата прибытия: {data['arrival_date']}\n"
            f"⏰ Время пребывания: {data['arrival_time']}\n"
            f"🏒 Команда/Школа: {data['team']}\n\n"
            "Администратор рассмотрит заявку и сообщит о решении.\n\n"
            "Для новой регистрации введите /start",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove(),
        )

        if ADMIN_CHAT_ID:
            admin_msg = (
                "🔔 *Новая заявка на въезд!*\n\n"
                f"🚗 {record['car_brand']}\n"
                f"🔢 `{record['car_number']}`\n"
                f"📅 Дата прибытия: {record['arrival_date']}\n"
                f"⏰ Время: {record['arrival_time']}\n"
                f"🏒 Команда/Школа: {record['team']}\n"
                f"📅 Зарегистрирован: {record['timestamp']}\n"
                f"🆔 TG: {record['tg_username']} (ID: {record['tg_id']})"
            )
            try:
                await context.bot.send_message(
                    chat_id=ADMIN_CHAT_ID,
                    text=admin_msg,
                    parse_mode="Markdown",
                )
            except Exception as e:
                logger.error(f"Не удалось отправить уведомление админу: {e}")
    else:
        await update.message.reply_text(
            "❌ Произошла ошибка при сохранении данных.\n"
            "Пожалуйста, попробуйте позже или обратитесь к администратору.",
            reply_markup=ReplyKeyboardRemove(),
        )

    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Отмена через команду /cancel."""
    context.user_data.clear()
    await update.message.reply_text(
        "🚫 Регистрация отменена.\n\nДля начала заново введите /start",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ConversationHandler.END


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Помощь."""
    await update.message.reply_text(
        "ℹ️ *Помощь*\n\n"
        "Этот бот регистрирует автомобиль для въезда на территорию.\n\n"
        "Команды:\n"
        "/start — Начать регистрацию\n"
        "/cancel — Отменить регистрацию\n"
        "/help — Эта справка",
        parse_mode="Markdown",
    )


def main():
    """Запуск бота."""
    app = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            CAR_BRAND: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_car_brand)],
            CAR_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_car_number)],
            ARRIVAL_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_arrival_date)],
            ARRIVAL_TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_arrival_time)],
            TEAM: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_team)],
            CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv_handler)
    app.add_handler(CommandHandler("help", help_command))

    logger.info("🤖 Бот запущен...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
