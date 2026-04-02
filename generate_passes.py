"""
Автоматическая генерация разовых пропусков и отправка в Telegram.
Запускается ежедневно в 23:30 МСК через GitHub Actions.
"""

import json
import logging
import os
import sys
import urllib.request
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

from num_to_words import number_to_genitive

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1ZGR8581dQu-rhvgnOBJ0AaIOqN6z-PKapZSa_k2qRkU")
SHEET_NAME = "Реестр автомобилей"
BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
RECIPIENT_CHAT_ID = os.environ.get("RECIPIENT_CHAT_ID", "5621135995")
TIMEZONE = ZoneInfo("Europe/Moscow")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)


def get_cars_for_date(target_date):
    """Получает из Google Sheets список авто на указанную дату."""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        logger.error("GOOGLE_CREDENTIALS не задана")
        return []
    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    records = sheet.get_all_records()
    cars = []
    for row in records:
        arrival = str(row.get("Дата прибытия", "")).strip()
        if arrival == target_date:
            brand = str(row.get("Марка", "")).strip()
            number = str(row.get("Гос. номер", "")).strip()
            if brand and number:
                cars.append({"car_brand": brand, "car_number": number})
    logger.info(f"Найдено {len(cars)} авто на {target_date}")
    return cars


def set_cell_border(cell):
    """Устанавливает границы ячейки таблицы."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = tcPr.makeelement(qn("w:tcBorders"), {})
        tcPr.append(tcBorders)
    for edge in ["top", "bottom", "left", "right"]:
        el = tcBorders.makeelement(
            qn(f"w:{edge}"),
            {qn("w:val"): "single", qn("w:sz"): "4", qn("w:space"): "0", qn("w:color"): "000000"},
        )
        tcBorders.append(el)


def generate_document(cars, target_date_str):
    """Генерирует Word-документ с разовыми пропусками."""
    target_date = datetime.strptime(target_date_str, "%d.%m.%Y")
    doc_date = datetime.now(TIMEZONE)
    doc_number = doc_date.strftime("%d-%m/%y")
    doc_date_fmt = doc_date.strftime("%d.%m.%Y")
    count = len(cars)
    count_words = number_to_genitive(count)

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    p = doc.add_paragraph()
    run = p.add_run(f"{doc_date_fmt}. \u2116 {doc_number}")
    run.font.size = Pt(12)

    p = doc.add_paragraph()
    run = p.add_run("\u041d\u0430 \u2116________ \u043e\u0442________")
    run.font.size = Pt(12)

    doc.add_paragraph()

    for text in ["\u0413\u0435\u043d\u0435\u0440\u0430\u043b\u044c\u043d\u043e\u043c\u0443 \u0434\u0438\u0440\u0435\u043a\u0442\u043e\u0440\u0443", "\u0413\u0411\u0423 \u00ab\u041c\u043e\u0441\u043f\u0440\u0438\u0440\u043e\u0434\u0430\u00bb"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(12)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("\u0410\u0434\u0438\u0433\u0430\u043c\u043e\u0432\u043e\u0439 \u042e.\u0418.")
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("\u0423\u0432\u0430\u0436\u0430\u0435\u043c\u0430\u044f \u042e\u043b\u0438\u044f \u0418\u043b\u044c\u0434\u0443\u0441\u043e\u0432\u043d\u0430!")
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph()
    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run("\u0412 \u0440\u0430\u043c\u043a\u0430\u0445 \u0440\u0430\u043d\u0435\u0435 \u043d\u0430\u043f\u0440\u0430\u0432\u043b\u0435\u043d\u043d\u043e\u0433\u043e \u043e\u0431\u0440\u0430\u0449\u0435\u043d\u0438\u044f \u0410\u041d\u041e \u0421\u041a \u00ab\u041e\u043b\u0438\u043c\u043f\u0438\u043a\u00bb \u043e\u0442 11.03.2026 \u2116 11-03/26,")
    run.font.size = Pt(12)

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(
        "\u0432 \u0441\u043e\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0438\u0438 \u0441 \u0434\u043e\u0433\u043e\u0432\u043e\u0440\u043e\u043c \u0430\u0440\u0435\u043d\u0434\u044b \u043e\u0431\u044a\u0435\u043a\u0442\u0430 \u043e\u0441\u043e\u0431\u043e \u0446\u0435\u043d\u043d\u043e\u0433\u043e \u0434\u0432\u0438\u0436\u0438\u043c\u043e\u0433\u043e \u0438\u043c\u0443\u0449\u0435\u0441\u0442\u0432\u0430 "
        "\u2116 \u041e\u0426\u0414\u0418-17, \u0437\u0430\u043a\u043b\u044e\u0447\u0435\u043d\u043d\u044b\u043c \u043c\u0435\u0436\u0434\u0443 \u0413\u041f\u0411\u0423 \u00ab\u041c\u043e\u0441\u043f\u0440\u0438\u0440\u043e\u0434\u0430\u00bb \u0438 \u0410\u0432\u0442\u043e\u043d\u043e\u043c\u043d\u043e\u0439 \u043d\u0435\u043a\u043e\u043c\u043c\u0435\u0440\u0447\u0435\u0441\u043a\u043e\u0439 "
        "\u041e\u0440\u0433\u0430\u043d\u0438\u0437\u0430\u0446\u0438\u0435\u0439 \u0421\u043f\u043e\u0440\u0442\u0438\u0432\u043d\u044b\u0439 \u043a\u043b\u0443\u0431 \u00ab\u041e\u043b\u0438\u043c\u043f\u0438\u043a\u00bb (\u041e\u0413\u0420\u041d 1047796891570) 01.08.2018 \u0433., "
        "\u0432 \u0446\u0435\u043b\u044f\u0445 \u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u044f \u0440\u0430\u0437\u043e\u0432\u044b\u0445 \u043e\u0431\u044f\u0437\u0430\u0442\u0435\u043b\u044c\u0441\u0442\u0432, \u043f\u0440\u043e\u0448\u0443 \u0412\u0430\u0441 \u0434\u0430\u0442\u044c \u0440\u0430\u0437\u043e\u0432\u044b\u0435 \u0440\u0430\u0437\u0440\u0435\u0448\u0435\u043d\u0438\u0435 \u043d\u0430 "
        "\u0432\u044a\u0435\u0437\u0434 \u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u043d\u044b\u0445 \u0441\u0440\u0435\u0434\u0441\u0442\u0432 \u043d\u0430 \u043e\u0441\u043e\u0431\u043e \u043e\u0445\u0440\u0430\u043d\u044f\u0435\u043c\u0443\u044e \u0437\u0435\u043b\u0435\u043d\u0443\u044e \u0442\u0435\u0440\u0440\u0438\u0442\u043e\u0440\u0438\u044e \u043b\u0430\u043d\u0434\u0448\u0430\u0444\u0442\u043d\u044b\u0439 "
        "\u0437\u0430\u043a\u0430\u0437\u043d\u0438\u043a \u00ab\u0422\u0435\u043f\u043b\u044b\u0439 \u0421\u0442\u0430\u043d\u00bb, \u043f\u043e\u0434\u0432\u0435\u0434\u043e\u043c\u0441\u0442\u0432\u0435\u043d\u043d\u0443\u044e \u0413\u0410\u0423\u041a \u0433. \u041c\u043e\u0441\u043a\u0432\u044b \u00ab\u041f\u0430\u0440\u043a\u0438 \u041c\u043e\u0441\u043a\u0432\u044b\u00bb, \u0434\u043b\u044f "
        "\u043f\u0440\u043e\u0435\u0437\u0434\u0430 \u043a \u043e\u0431\u044a\u0435\u043a\u0442\u0430\u043c, \u0440\u0430\u0441\u043f\u043e\u043b\u043e\u0436\u0435\u043d\u043d\u044b\u043c \u043f\u043e \u0430\u0434\u0440\u0435\u0441\u0443: \u0433. \u041c\u043e\u0441\u043a\u0432\u0430, \u0443\u043b. \u041e\u0441\u0442\u0440\u043e\u0432\u0438\u0442\u044f\u043d\u043e\u0432\u0430, \u0432\u043b.10., "
        f"\u043d\u0430 {target_date_str} \u0433., \u0441\u043e\u0433\u043b\u0430\u0441\u043d\u043e \u043f\u0440\u0438\u043b\u0430\u0433\u0430\u0435\u043c\u043e\u0439 \u0441\u0445\u0435\u043c\u0435 \u0434\u0432\u0438\u0436\u0435\u043d\u0438\u044f \u0430\u0432\u0442\u043e\u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u0430."
    )
    run.font.size = Pt(12)

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(1.27)
    run = p.add_run("\u0418\u043d\u0444\u043e\u0440\u043c\u0430\u0446\u0438\u044f \u043e \u043c\u0430\u0440\u043a\u0430\u0445 \u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u043d\u044b\u0445 \u0441\u0440\u0435\u0434\u0441\u0442\u0432 \u0438 \u0433\u043e\u0441\u0443\u0434\u0430\u0440\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0445 \u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0430\u0446\u0438\u043e\u043d\u043d\u044b\u0445 \u0437\u043d\u0430\u043a\u0430\u0445:")
    run.font.size = Pt(12)

    doc.add_paragraph()

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, h in enumerate(["\u2116", "\u041d\u043e\u043c\u0435\u0440 \u0430\u0432\u0442\u043e\u043c\u043e\u0431\u0438\u043b\u044f", "\u041c\u0430\u0440\u043a\u0430 \u0430\u0432\u0442\u043e\u043c\u043e\u0431\u0438\u043b\u044f"]):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(11)
        set_cell_border(cell)

    for idx, car in enumerate(cars, 1):
        row = table.add_row()
        for i, val in enumerate([str(idx), car["car_brand"], car["car_number"]]):
            cell = row.cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(val)
            run.font.size = Pt(11)
            run.font.name = "Arial"
            set_cell_border(cell)

    for row in table.rows:
        row.cells[0].width = Cm(1.5)
        row.cells[1].width = Cm(7)
        row.cells[2].width = Cm(7)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(
        f"\u0422\u0430\u043a\u0436\u0435 \u0441\u043e\u043e\u0431\u0449\u0430\u044e, \u0447\u0442\u043e \u0435\u0434\u0438\u043d\u043e\u0432\u0440\u0435\u043c\u0435\u043d\u043d\u043e \u043d\u0430 \u0442\u0435\u0440\u0440\u0438\u0442\u043e\u0440\u0438\u0438 \u043b\u0430\u043d\u0434\u0448\u0430\u0444\u0442\u043d\u043e\u0433\u043e \u0437\u0430\u043a\u0430\u0437\u043d\u0438\u043a\u0430 "
        f"\u00ab\u0422\u0435\u043f\u043b\u044b\u0439 \u0421\u0442\u0430\u043d\u00bb \u043f\u043e \u0430\u0434\u0440\u0435\u0441\u0443 : \u0433. \u041c\u043e\u0441\u043a\u0432\u0430 \u0443\u043b. \u041e\u0441\u0442\u0440\u043e\u0432\u0438\u0442\u044f\u043d\u043e\u0432\u0430, \u0432\u043b.10, \u0437\u0430\u0435\u0437\u0434 "
        f"\u0430\u0432\u0442\u043e\u043c\u043e\u0431\u0438\u043b\u0435\u0439 \u043d\u0435 \u0431\u0443\u0434\u0435\u0442 \u043f\u0440\u0435\u0432\u044b\u0448\u0430\u0442\u044c \u043a\u043e\u043b\u0438\u0447\u0435\u0441\u0442\u0432\u0430 {count} ({count_words}) \u0448\u0442\u0443\u043a, "
        f"\u0430 \u0442\u0430\u043a\u0436\u0435 \u0431\u0443\u0434\u0443\u0442 \u0441\u043e\u0431\u043b\u044e\u0434\u0430\u0442\u044c\u0441\u044f \u0440\u0435\u0433\u043b\u0430\u043c\u0435\u043d\u0442"
    )
    run.font.size = Pt(12)

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(
        "\u043f\u043e\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u044f \u041f\u0440\u0430\u0432\u0438\u0442\u0435\u043b\u044c\u0441\u0442\u0432\u0430 \u041c\u043e\u0441\u043a\u0432\u044b \u043e\u0442 14.09.2010 \u2116795-\u041f\u041f \u00ab\u041e\u0431 \u0443\u0442\u0432\u0435\u0440\u0436\u0434\u0435\u043d\u0438\u0438 "
        "\u0420\u0435\u0433\u043b\u0430\u043c\u0435\u043d\u0442\u0430 \u043f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u043a\u0438 \u0438 \u0432\u044b\u0434\u0430\u0447\u0438 \u0437\u0430\u044f\u0432\u0438\u0442\u0435\u043b\u0435\u043c \u0414\u0435\u043f\u0430\u0440\u0442\u0430\u043c\u0435\u043d\u0442\u043e\u043c \u043f\u0440\u0438\u0440\u043e\u0434\u043e\u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u043d\u0438\u044f "
        "\u0438 \u043e\u0445\u0440\u0430\u043d\u044b \u043e\u043a\u0440\u0443\u0436\u0430\u044e\u0449\u0435\u0439 \u0441\u0440\u0435\u0434\u044b \u0433\u043e\u0440\u043e\u0434\u0430 \u041c\u043e\u0441\u043a\u0432\u044b \u043d\u0430 \u0432\u044a\u0435\u0437\u0434 \u043d\u0430 \u043e\u0441\u043e\u0431\u043e \u043e\u0445\u0440\u0430\u043d\u044f\u0435\u043c\u044b\u0435 \u043f\u0440\u0438\u0440\u043e\u0434\u043d\u044b\u0435 "
        "\u0442\u0435\u0440\u0440\u0438\u0442\u043e\u0440\u0438\u0438 \u0433\u043e\u0440\u043e\u0434\u0430 \u041c\u043e\u0441\u043a\u0432\u044b\u00bb, \u0430 \u0438\u043c\u0435\u043d\u043d\u043e:"
    )
    run.font.size = Pt(12)

    for bullet in [
        "\u0432\u044a\u0435\u0437\u0434 \u0438 \u043f\u0435\u0440\u0435\u0434\u0432\u0438\u0436\u0435\u043d\u0438\u0435 \u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u043d\u044b\u0445 \u0441\u0440\u0435\u0434\u0441\u0442\u0432 \u043f\u043e \u041e\u041e\u0417\u0422 \u0432\u043d\u0435 \u0434\u043e\u0440\u043e\u0433 \u043e\u0431\u0449\u0435\u0433\u043e \u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u043d\u0438\u044f \u0431\u0443\u0434\u0435\u0442 \u043e\u0441\u0443\u0449\u0435\u0441\u0442\u0432\u043b\u044f\u0442\u044c\u0441\u044f \u043f\u043e \u0441\u0442\u0440\u043e\u0433\u043e \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043d\u044b\u043c \u043c\u0430\u0440\u0448\u0440\u0443\u0442\u0430\u043c (\u0442\u0435\u0445\u043d\u043e\u043b\u043e\u0433\u0438\u0447\u0435\u0441\u043a\u0438\u043c \u043a\u0430\u0440\u0442\u0430\u043c), \u0441\u043e\u0433\u043b\u0430\u0441\u043e\u0432\u0430\u043d\u043d\u044b\u043c \u0431\u0430\u043b\u0430\u043d\u0441\u043e\u0434\u0435\u0440\u0436\u0430\u0442\u0435\u043b\u0435\u043c \u0442\u0435\u0440\u0440\u0438\u0442\u043e\u0440\u0438\u0438 \u0413\u0410\u0423\u041a \u0433. \u041c\u043e\u0441\u043a\u0432\u044b \u00ab\u041f\u0430\u0440\u043a\u0438 \u041c\u043e\u0441\u043a\u0432\u044b\u00bb \u0438 \u0413\u0411\u0423 \u00ab\u041c\u043e\u0441\u043f\u0440\u0438\u0440\u043e\u0434\u0430\u00bb;",
        "\u0432\u044a\u0435\u0437\u0434, \u043f\u0435\u0440\u0435\u0434\u0432\u0438\u0436\u0435\u043d\u0438\u0435 \u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u043d\u044b\u0445 \u0441\u0440\u0435\u0434\u0441\u0442\u0432 \u043f\u043e \u041e\u041e\u0417\u0422 \u0432\u043d\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043d\u044b\u0445 \u043c\u0430\u0440\u0448\u0440\u0443\u0442\u043e\u0432, \u0430 \u0442\u0430\u043a \u0436\u0435 \u043e\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u0438 \u0441\u0442\u043e\u044f\u043d\u043a\u0430 \u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u043d\u044b\u0445 \u0441\u0440\u0435\u0434\u0441\u0442\u0432 \u0432\u043d\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043d\u044b\u0445 \u043c\u0435\u0441\u0442 \u0431\u0443\u0434\u0435\u0442 \u0437\u0430\u043f\u0440\u0435\u0449\u0435\u043d\u0430;",
        "\u041f\u0440\u043e\u0435\u0437\u0434 \u0431\u0443\u0434\u0435\u0442 \u043e\u0441\u0443\u0449\u0435\u0441\u0442\u0432\u043b\u044f\u0442\u044c\u0441\u044f \u043f\u043e \u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u044e\u0449\u0435\u0439 \u0434\u043e\u0440\u043e\u0436\u043d\u043e\u0439 \u0441\u0435\u0442\u0438 \u0441 \u0438\u0441\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435\u043c \u0437\u0430\u0435\u0437\u0434\u0430 \u0430\u0432\u0442\u043e\u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u0430 \u043d\u0430 \u0442\u0435\u0440\u0440\u0438\u0442\u043e\u0440\u0438\u0438, \u0437\u0430\u043d\u044f\u0442\u044b\u0435 \u0437\u0435\u043b\u0435\u043d\u044b\u043c\u0438 \u043d\u0430\u0441\u0430\u0436\u0434\u0435\u043d\u0438\u044f\u043c\u0438.",
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.add_run(f"- {bullet}")
        run.font.size = Pt(12)

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(
        "\u0422\u0430\u043a\u0436\u0435 \u0431\u0443\u0434\u0435\u0442 \u043e\u0441\u0443\u0449\u0435\u0441\u0442\u0432\u043b\u0435\u043d \u0432\u0438\u0434\u0435\u043e\u043a\u043e\u043d\u0442\u0440\u043e\u043b\u044c \u0430\u0432\u0442\u043e\u043c\u043e\u0431\u0438\u043b\u044c\u043d\u044b\u0445 \u043d\u043e\u043c\u0435\u0440\u043e\u0432 \u0432\u044a\u0435\u0437\u0436\u0430\u044e\u0449\u0435\u0433\u043e "
        "\u043d\u0430 \u0442\u0435\u0440\u0440\u0438\u0442\u043e\u0440\u0438\u044e \u043b\u0430\u043d\u0434\u0448\u0430\u0444\u0442\u043d\u043e\u0433\u043e \u0437\u0430\u043a\u0430\u0437\u043d\u0438\u043a\u0430 \u00ab\u0422\u0435\u043f\u043b\u044b\u0439 \u0421\u0442\u0430\u043d\u00bb \u0430\u0432\u0442\u043e\u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u0430 \u043f\u043e \u0430\u0434\u0440\u0435\u0441\u0443: "
        "\u0433. \u041c\u043e\u0441\u043a\u0432\u0430, \u0443\u043b. \u041e\u0441\u0442\u0440\u043e\u0432\u0438\u0442\u044f\u043d\u043e\u0432\u0430, \u0432\u043b.10."
    )
    run.font.size = Pt(12)

    doc.add_paragraph()
    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run("\u041f\u0440\u0435\u0434\u0441\u0435\u0434\u0430\u0442\u0435\u043b\u044c")
    run.font.size = Pt(12)
    p = doc.add_paragraph()
    run = p.add_run("\u041f\u0440\u0430\u0432\u043b\u0435\u043d\u0438\u044f \u0410\u041d\u041e \u0421\u041a \u00ab\u041e\u041b\u0418\u041c\u041f\u0418\u041a\u00bb")
    run.font.size = Pt(12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("\u0410.\u041f. \u0414\u043c\u0438\u0442\u0440\u0438\u0435\u043d\u043a\u043e")
    run.font.size = Pt(12)

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run("\u0418\u0441\u043f\u043e\u043b\u043d\u0438\u0442\u0435\u043b\u044c")
    run.font.size = Pt(12)
    p = doc.add_paragraph()
    run = p.add_run("\u041f\u0440\u0443\u0446\u043a\u043e\u0432 \u0418\u043b\u044c\u044f \u0415\u0432\u0433\u0435\u043d\u044c\u0435\u0432\u0438\u0447")
    run.font.size = Pt(12)
    p = doc.add_paragraph()
    run = p.add_run("+79166571350")
    run.font.size = Pt(12)

    date_fn = target_date.strftime("%d_%m_%Y")
    filename = f"Razovye_propuska_{date_fn}.docx"
    filepath = Path("/tmp") / filename
    doc.save(str(filepath))
    logger.info(f"\u0414\u043e\u043a\u0443\u043c\u0435\u043d\u0442 \u0441\u043e\u0445\u0440\u0430\u043d\u0451\u043d: {filepath}")
    return filepath


def send_telegram_document(filepath, caption):
    """Отправляет файл через Telegram Bot API."""
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    boundary = "----FormBoundary7MA4YWxkTrZu0gW"
    filename = filepath.name
    with open(filepath, "rb") as f:
        file_data = f.read()
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="chat_id"\r\n\r\n'
        f"{RECIPIENT_CHAT_ID}\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="caption"\r\n\r\n'
        f"{caption}\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="document"; filename="{filename}"\r\n'
        f"Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    ).encode("utf-8") + file_data + f"\r\n--{boundary}--\r\n".encode("utf-8")
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    resp = urllib.request.urlopen(req)
    result = json.loads(resp.read())
    if result.get("ok"):
        logger.info(f"\u0424\u0430\u0439\u043b \u043e\u0442\u043f\u0440\u0430\u0432\u043b\u0435\u043d \u0432 Telegram: {filename}")
    else:
        logger.error(f"\u041e\u0448\u0438\u0431\u043a\u0430 \u043e\u0442\u043f\u0440\u0430\u0432\u043a\u0438: {result}")
        sys.exit(1)


def send_telegram_message(text):
    """Отправляет текстовое сообщение через Telegram Bot API."""
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    data = json.dumps({"chat_id": RECIPIENT_CHAT_ID, "text": text}).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"}, method="POST")
    resp = urllib.request.urlopen(req)
    result = json.loads(resp.read())
    if result.get("ok"):
        logger.info("\u0421\u043e\u043e\u0431\u0449\u0435\u043d\u0438\u0435 \u043e\u0442\u043f\u0440\u0430\u0432\u043b\u0435\u043d\u043e \u0432 Telegram")
    else:
        logger.error(f"\u041e\u0448\u0438\u0431\u043a\u0430 \u043e\u0442\u043f\u0440\u0430\u0432\u043a\u0438: {result}")


def main():
    if not BOT_TOKEN:
        logger.error("BOT_TOKEN \u043d\u0435 \u0437\u0430\u0434\u0430\u043d")
        sys.exit(1)

    now = datetime.now(TIMEZONE)
    tomorrow = now + timedelta(days=1)
    target_date_str = tomorrow.strftime("%d.%m.%Y")

    logger.info(f"\u0413\u0435\u043d\u0435\u0440\u0430\u0446\u0438\u044f \u043f\u0440\u043e\u043f\u0443\u0441\u043a\u043e\u0432 \u043d\u0430 {target_date_str}")
    cars = get_cars_for_date(target_date_str)

    if not cars:
        msg = f"\u041d\u0430 {target_date_str} \u0437\u0430\u044f\u0432\u043e\u043a \u043d\u0430 \u043f\u0440\u043e\u043f\u0443\u0441\u043a\u0430 \u043d\u0435\u0442."
        logger.info(msg)
        send_telegram_message(msg)
        return

    filepath = generate_document(cars, target_date_str)
    caption = f"\u0420\u0430\u0437\u043e\u0432\u044b\u0435 \u043f\u0440\u043e\u043f\u0443\u0441\u043a\u0430 \u043d\u0430 {target_date_str}\n\u0410\u0432\u0442\u043e\u043c\u043e\u0431\u0438\u043b\u0435\u0439: {len(cars)}"
    send_telegram_document(filepath, caption)
    filepath.unlink(missing_ok=True)
    logger.info("\u0413\u043e\u0442\u043e\u0432\u043e!")


if __name__ == "__main__":
    main()
