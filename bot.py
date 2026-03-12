"""
Telegram бот для регистрации автомобилей на территорию.
Данные сохраняются в Google Таблицы.
"""

import logging
import re
from datetime import datetime
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
PHONE, CAR_BRAND, CAR_MODEL, CAR_NUMBER, CONFIRM = range(5)

# Кнопки подтверждения
CONFIRM_KEYBOARD = ReplyKeyboardMarkup(
    [["✅ Подтвердить", "❌ Отменить"]],
    resize_keyboard=True,
    one_time_keyboard=True,
)


def format_car_number(number: str) -> str:
    """Приводит номер к верхнему регистру и убирает лишние пробелы."""
    return number.strip().upper()


def is_valid_phone(phone: str) -> bool:
    """Проверяет базовый формат телефонного номера."""
    cleaned = re.sub(r"[\s\-\(\)]", "", phone)
    return bool(re.match(r"^[\+7|8]?\d{10,11}$", cleaned))


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Начало диалога регистрации."""
    user = update.effective_user
    logger.info(f"Пользователь {user.id} ({user.username}) начал регистрацию")

    await update.message.reply_text(
        f"👋 Добро пожаловать, {user.first_name}!\n\n"
        "🚗 Этот бот поможет зарегистрировать ваш автомобиль для въезда на территорию.\n\n"
        "Пожалуйста, заполните данные. Это займёт около минуты.\n\n"
        "📱 *Шаг 1 из 4* — Введите номер телефона:",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove(),
    )
    return PHONE


async def get_phone(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение телефона."""
    phone = update.message.text.strip()

    if not is_valid_phone(phone):
        await update.message.reply_text(
            "⚠️ Неверный формат номера. Введите номер в формате:\n"
            "+7XXXXXXXXXX или 8XXXXXXXXXX"
        )
        return PHONE

    context.user_data["phone"] = phone
    await update.message.reply_text(
        f"✅ Телефон: *{phone}*\n\n🚗 *Шаг 2 из 4* — Введите марку автомобиля:\n"
        "_Пример: Toyota, BMW, Lada_",
        parse_mode="Markdown",
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
        f"✅ Марка: *{brand.title()}*\n\n🚗 *Шаг 3 из 4* — Введите модель автомобиля:\n"
        "_Пример: Camry, X5, Vesta_",
        parse_mode="Markdown",
    )
    return CAR_MODEL


async def get_car_model(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение модели авто."""
    model = update.message.text.strip()

    if len(model) < 1:
        await update.message.reply_text("⚠️ Введите корректное название модели:")
        return CAR_MODEL

    context.user_data["car_model"] = model.title()
    await update.message.reply_text(
        f"✅ Модель: *{model.title()}*\n\n🔢 *Шаг 4 из 4* — Введите гос. номер автомобиля:\n"
        "_Пример: А123БВ777_",
        parse_mode="Markdown",
    )
    return CAR_NUMBER


async def get_car_number(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение гос. номера и показ итоговой формы."""
    number = format_car_number(update.message.text)

    if len(number) < 4:
        await update.message.reply_text(
            "⚠️ Введите корректный гос. номер автомобиля:"
        )
        return CAR_NUMBER

    context.user_data["car_number"] = number

    # Итоговая сводка
    data = context.user_data
    summary = (
        "📋 *Проверьте введённые данные:*\n\n"
        f"📱 Телефон: {data['phone']}\n"
        f"🚗 Автомобиль: {data['car_brand']} {data['car_model']}\n"
        f"🔢 Гос. номер: `{data['car_number']}`\n\n"
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

    # Сохраняем данные
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

    # Запись в таблицу
    record = {
        "timestamp": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "tg_id": str(user.id),
        "tg_username": f"@{user.username}" if user.username else "—",
        "phone": data["phone"],
        "car_brand": data["car_brand"],
        "car_model": data["car_model"],
        "car_number": data["car_number"],
        "status": "Ожидает одобрения",
    }

    success = sheets.add_record(record)

    if success:
        await update.message.reply_text(
            "✅ *Заявка успешно отправлена!*\n\n"
            f"🔢 Ваш номер авто: `{data['car_number']}`\n\n"
            "Администратор рассмотрит заявку и сообщит о решении.\n\n"
            "Для новой регистрации введите /start",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove(),
        )

        # Уведомление администратору
        if ADMIN_CHAT_ID:
            admin_msg = (
                "🔔 *Новая заявка на въезд!*\n\n"
                f"📱 {record['phone']}\n"
                f"🚗 {record['car_brand']} {record['car_model']}\n"
                f"🔢 `{record['car_number']}`\n"
                f"📅 {record['timestamp']}\n"
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

    # Диалог регистрации
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_phone)],
            CAR_BRAND: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_car_brand)],
            CAR_MODEL: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_car_model)],
            CAR_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_car_number)],
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
