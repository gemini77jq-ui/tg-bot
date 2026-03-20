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

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Заголовки таблицы (порядок важен!)
HEADERS = [
    "Дата регистрации",
    "TG ID",
    "TG Username",
    "Марка",
    "Гос. номер",
    "Дата прибытия",
    "Время пребывания",
    "Команда/Школа",
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

            try:
                self._sheet = spreadsheet.worksheet(SHEET_NAME)
            except gspread.WorksheetNotFound:
                self._sheet = spreadsheet.add_worksheet(
                    title=SHEET_NAME, rows=1000, cols=len(HEADERS)
                )
                self._setup_headers()
                return True

            # Обновляем заголовки если не совпадают
            existing = self._sheet.row_values(1)
            if existing != HEADERS:
                self._setup_headers()

            return True
        except Exception as e:
            logger.error(f"Ошибка подключения к Google Sheets: {e}")
            return False

    def _setup_headers(self):
        """Устанавливает заголовки таблицы."""
        try:
            self._sheet.update("A1", [HEADERS])
            self._sheet.format(
                f"A1:{chr(64 + len(HEADERS))}1",
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
        timestamp, tg_id, tg_username, car_brand, car_number,
        arrival_date, arrival_time, team, status
        """
        if not self._connect():
            return False

        try:
            row = [
                record["timestamp"],
                record["tg_id"],
                record["tg_username"],
                record["car_brand"],
                record["car_number"],
                record["arrival_date"],
                record["arrival_time"],
                record["team"],
                record["status"],
            ]
            self._sheet.append_row(row, value_input_option="USER_ENTERED")
            logger.info(f"Запись добавлена: {record['car_number']}")
            return True
        except Exception as e:
            logger.error(f"Ошибка добавления записи: {e}")
            return False

    def is_duplicate(self, car_number: str) -> bool:
        """
        Проверяет, зарегистрирован ли уже автомобиль с таким номером.
        Гос. номер — столбец E (индекс 5).
        """
        if not self._connect():
            return False

        try:
            all_numbers = self._sheet.col_values(5)  # столбец E
            registered = [n.strip().upper() for n in all_numbers[1:] if n]
            return car_number.upper() in registered
        except Exception as e:
            logger.error(f"Ошибка проверки дубликата: {e}")
            return False

    def get_all_records(self) -> list:
        """Возвращает все записи из таблицы."""
        if not self._connect():
            return []

        try:
            return self._sheet.get_all_records()
        except Exception as e:
            logger.error(f"Ошибка получения записей: {e}")
            return []
