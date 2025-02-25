import asyncio
import os
import datetime
from docx import template
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, ConversationHandler, CallbackContext
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.filters import Command
from aiogram.enums import ChatType
from config import Config, load_config

# Настройка логирования
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Загрузка конфигурации бота
config: Config = load_config()
BOT_TOKEN: str = config.tg_bot.token

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()


# Константы для состояний
GET_CUSTOMER_NAME, GET_CONTRACT_AMOUNT, GET_PRODUCT_NAME, GET_BANK_DETAILS = range(4)

# Шаблон договора
TEMPLATE_PATH = "template.docx"

# Функция для замены меток в документе
def replace_placeholders(doc, placeholders):
    for paragraph in doc.paragraphs:
        for key, value in placeholders.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
                # Выравнивание первой даты по центру
                if key == "{Сегодняшняя дата}":
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # Выравнивание ФИО заказчика по центру
                if key == "{Заказчик}":
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Замена в таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in placeholders.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

# Функция для создания PDF из DOCX
def create_pdf(docx_path, pdf_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Times", size=13)
    
    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        pdf.multi_cell(0, 10, paragraph.text)
    
    pdf.output(pdf_path)

# Обработчик команды /start
def start(update: Update, context: CallbackContext) -> int:
    update.message.reply_text("Введите ФИО заказчика:")
    return GET_CUSTOMER_NAME

# Обработчик ввода ФИО заказчика
def get_customer_name(update: Update, context: CallbackContext) -> int:
    context.user_data['customer_name'] = update.message.text
    update.message.reply_text("Введите сумму договора (цифрами):")
    return GET_CONTRACT_AMOUNT

# Обработчик ввода суммы договора
def get_contract_amount(update: Update, context: CallbackContext) -> int:
    context.user_data['contract_amount'] = update.message.text
    update.message.reply_text("Введите название товара в родительном падеже:")
    return GET_PRODUCT_NAME

# Обработчик ввода названия товара
def get_product_name(update: Update, context: CallbackContext) -> int:
    context.user_data['product_name'] = update.message.text
    update.message.reply_text("Введите банковские реквизиты (ИНН, ОГРНИП, расчетный счет, банк, БИК, корр. счет, телефон):")
    return GET_BANK_DETAILS

# Обработчик ввода банковских реквизитов
def get_bank_details(update: Update, context: CallbackContext) -> int:
    context.user_data['bank_details'] = update.message.text

    # Заполнение шаблона
    doc = Document(TEMPLATE_PATH)
    placeholders = {
        "{Заказчик}": f"Индивидуальный Предприниматель {context.user_data['customer_name']}",
        "{Сегодняшняя дата}": datetime.datetime.now().strftime("%d.%m.%Y"),
        "{Название товара в родительном падеже}": context.user_data['product_name'],
        "{Стоимость работ цифрами}": context.user_data['contract_amount'],
        "{Стоимость работ прописью}": num2words(int(context.user_data['contract_amount']), lang='ru') + " рублей 00 копеек",
        "{Банковские реквизиты}": context.user_data['bank_details']
    }
    replace_placeholders(doc, placeholders)

    # Сохранение DOCX
    docx_output_path = "output.docx"
    doc.save(docx_output_path)

    # Создание PDF
    pdf_output_path = "output.pdf"
    create_pdf(docx_output_path, pdf_output_path)

    # Отправка файлов пользователю
    update.message.reply_document(open(docx_output_path, 'rb'))
    update.message.reply_document(open(pdf_output_path, 'rb'))

    return ConversationHandler.END

# Основная функция
def main():
    # Токен бота
    token = "YOUR_TELEGRAM_BOT_TOKEN"
    updater = Updater(token, use_context=True)
    dp = updater.dispatcher

    # Диалог
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            GET_CUSTOMER_NAME: [MessageHandler(Filters.text & ~Filters.command, get_customer_name)],
            GET_CONTRACT_AMOUNT: [MessageHandler(Filters.text & ~Filters.command, get_contract_amount)],
            GET_PRODUCT_NAME: [MessageHandler(Filters.text & ~Filters.command, get_product_name)],
            GET_BANK_DETAILS: [MessageHandler(Filters.text & ~Filters.command, get_bank_details)],
        },
        fallbacks=[]
    )

    dp.add_handler(conv_handler)
    updater.start_polling()
    updater.idle()

if __name__ == "__main__":
    main()