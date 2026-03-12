"""
Менеджер Google Таблиц.
Отвечает за запись данных и проверку дубликатов.
"""

import logging
import json
import os
from typing import Optional
import gspread
from google.oauth2.service_account import Credentials
from config import SPREADSHEET_ID, SHEET_NAME

logger = logging.getLogger(__name__)

# Права доступа к Google API
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Заголовки таблицы (порядок важен!)
HEADERS = [
    "Дата и время",
    "TG ID",
    "TG Username",
    "Телефон",
    "Марка",
    "Модель",
    "Гос. номер",
    "Статус",
]


class GoogleSheetsManager:
    def __init__(self):
        self._client: Optional[gspread.Client] = None
        self._sheet: Optional[gspread.Worksheet] = None

    def _connect(self) -> bool:
        """Подключение к Google Sheets."""
        try:
            credentials_json = os.environ.get("GOOGLE_CREDENTIALS")
            if not credentials_json:
                logger.error("Переменная GOOGLE_CREDENTIALS не найдена")
                return False

            credentials_dict = json.loads(credentials_json)
            creds = Credentials.from_service_account_info(
                credentials_dict, scopes=SCOPES
            )
            self._client = gspread.authorize(creds)
            spreadsheet = self._client.open_by_key(SPREADSHEET_ID)

            # Получаем лист или создаём новый
            try:
                self._sheet = spreadsheet.worksheet(SHEET_NAME)
            except gspread.WorksheetNotFound:
                self._sheet = spreadsheet.add_worksheet(
                    title=SHEET_NAME, rows=1000, cols=len(HEADERS)
                )
                self._setup_headers()

            # Проверяем/создаём заголовки
            existing = self._sheet.row_values(1)
            if not existing:
                self._setup_headers()

            return True
        except Exception as e:
            logger.error(f"Ошибка подключения к Google Sheets: {e}")
            return False

    def _setup_headers(self):
        """Устанавливает заголовки таблицы."""
        try:
            self._sheet.update("A1", [HEADERS])
            # Форматирование заголовков (жирный)
            self._sheet.format(
                "A1:H1",
                {
                    "textFormat": {"bold": True},
                    "backgroundColor": {"red": 0.2, "green": 0.6, "blue": 1.0},
                },
            )
            logger.info("Заголовки таблицы установлены")
        except Exception as e:
            logger.error(f"Ошибка установки заголовков: {e}")

    def add_record(self, record: dict) -> bool:
        """
        Добавляет запись в таблицу.
        
        record должен содержать ключи:
        timestamp, tg_id, tg_username, full_name, phone,
        car_brand, car_model, car_number, status
        """
        if not self._connect():
            return False

        try:
            row = [
                record["timestamp"],
                record["tg_id"],
                record["tg_username"],
                record["phone"],
                record["car_brand"],
                record["car_model"],
                record["car_number"],
                record["status"],
            ]
            self._sheet.append_row(row, value_input_option="USER_ENTERED")
            logger.info(f"Запись добавлена: {record['car_number']} — {record['full_name']}")
            return True
        except Exception as e:
            logger.error(f"Ошибка добавления записи: {e}")
            return False

    def is_duplicate(self, car_number: str) -> bool:
        """
        Проверяет, зарегистрирован ли уже автомобиль с таким номером.
        Столбец 'Гос. номер' — индекс 8 (H).
        """
        if not self._connect():
            return False

        try:
            # Получаем все номера из столбца H (индекс 8)
            all_numbers = self._sheet.col_values(7)  # 1-based индекс
            # Пропускаем заголовок и нормализуем
            registered = [n.strip().upper() for n in all_numbers[1:] if n]
            return car_number.upper() in registered
        except Exception as e:
            logger.error(f"Ошибка проверки дубликата: {e}")
            return False  # При ошибке разрешаем регистрацию

    def get_all_records(self) -> list:
        """Возвращает все записи из таблицы (для отладки/admin)."""
        if not self._connect():
            return []

        try:
            return self._sheet.get_all_records()
        except Exception as e:
            logger.error(f"Ошибка получения записей: {e}")
            return []
