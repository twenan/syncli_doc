import asyncio
import os
import datetime
import logging
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from fpdf import FPDF
from num2words import num2words
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.utils.keyboard import ReplyKeyboardBuilder
from config import Config, load_config

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Загрузка конфигурации бота
config: Config = load_config()
BOT_TOKEN: str = config.tg_bot.token

# Создание бота и диспетчера
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# Определение состояний для FSM
class ContractStates(StatesGroup):
    GET_CUSTOMER_NAME = State()
    GET_CONTRACT_AMOUNT = State()
    GET_PRODUCT_NAME = State()
    GET_BANK_DETAILS = State()

# Шаблон договора
TEMPLATE_PATH = "template.docx"

# Функция для замены меток в документе
def replace_placeholders(doc, placeholders):
    for paragraph in doc.paragraphs:
        for key, value in placeholders.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
                if key == "{Сегодняшняя дата}":
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if key == "{Заказчик}":
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

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

@dp.message(Command("start"))
async def start(message: types.Message, state: FSMContext):
    welcome_text = (
        "🤖 **Что умеет этот бот?**\n"
        "Этот бот предназначен для помощи в заполнении договоров.\n\n"
        "Готовы начать? Нажмите кнопку ниже 👇"
    )

    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.add(types.KeyboardButton("🚀 Начать заполнение договора"))

    await message.answer(welcome_text, reply_markup=keyboard, parse_mode="Markdown")
    await state.set_state(ContractStates.GET_CUSTOMER_NAME)


# Обработчик ввода ФИО заказчика
@dp.message(ContractStates.GET_CUSTOMER_NAME)
async def get_customer_name(message: types.Message, state: FSMContext):
    await state.update_data(customer_name=message.text)
    await message.answer("Введите сумму договора (цифрами):")
    await state.set_state(ContractStates.GET_CONTRACT_AMOUNT)

# Обработчик ввода суммы договора
@dp.message(ContractStates.GET_CONTRACT_AMOUNT)
async def get_contract_amount(message: types.Message, state: FSMContext):
    await state.update_data(contract_amount=message.text)
    await message.answer("Введите название товара в родительном падеже:")
    await state.set_state(ContractStates.GET_PRODUCT_NAME)

# Обработчик ввода названия товара
@dp.message(ContractStates.GET_PRODUCT_NAME)
async def get_product_name(message: types.Message, state: FSMContext):
    await state.update_data(product_name=message.text)
    await message.answer("Введите банковские реквизиты (ИНН, ОГРНИП, расчетный счет, банк, БИК, корр. счет, телефон):")
    await state.set_state(ContractStates.GET_BANK_DETAILS)

# Обработчик ввода банковских реквизитов
@dp.message(ContractStates.GET_BANK_DETAILS)
async def get_bank_details(message: types.Message, state: FSMContext):
    await state.update_data(bank_details=message.text)
    data = await state.get_data()

    # Заполнение шаблона
    doc = Document(TEMPLATE_PATH)
    placeholders = {
        "{Заказчик}": f"Индивидуальный Предприниматель {data['customer_name']}",
        "{Сегодняшняя дата}": datetime.datetime.now().strftime("%d.%m.%Y"),
        "{Название товара в родительном падеже}": data['product_name'],
        "{Стоимость работ цифрами}": data['contract_amount'],
        "{Стоимость работ прописью}": num2words(int(data['contract_amount']), lang='ru') + " рублей 00 копеек",
        "{Банковские реквизиты}": data['bank_details']
    }
    replace_placeholders(doc, placeholders)

    # Сохранение DOCX
    docx_output_path = "output.docx"
    doc.save(docx_output_path)

    # Создание PDF
    pdf_output_path = "output.pdf"
    create_pdf(docx_output_path, pdf_output_path)

    # Проверим, созданы ли файлы
    if os.path.exists(docx_output_path) and os.path.exists(pdf_output_path):
        # Отправка документов пользователю
        try:
            await message.answer("Готовый договор отправлен:")
            await message.answer_document(types.FSInputFile(docx_output_path))
            await message.answer_document(types.FSInputFile(pdf_output_path))
        except Exception as e:
            await message.answer(f"Ошибка при отправке файла: {str(e)}")
    else:
        await message.answer("Ошибка: не удалось создать файлы договора.")

    await state.clear()

# Основная функция запуска бота
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
