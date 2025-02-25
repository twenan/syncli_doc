import asyncio
import os
import datetime
import logging
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt  # Для установки размера шрифта
from docx.oxml.ns import qn  # Для поддержки русских символов
from fpdf import FPDF
from datetime import datetime, timedelta
from num2words import num2words
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.storage.memory import MemoryStorage
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
        full_text = ''.join(run.text for run in paragraph.runs)
        logging.info(f"Текст параграфа до замены: {full_text}")

        for key, value in placeholders.items():
            logging.info(f"Проверяем плейсхолдер: '{key}'")
            if key.lower() in full_text.lower():
                logging.info(f"Плейсхолдер '{key}' найден в параграфе!")

                for run in paragraph.runs:
                    if key.lower() in run.text.lower():
                        original_text = run.text
                        updated_text = replace_with_case_preservation(original_text, key, value)
                        logging.info(f"Заменяем '{original_text}' на '{updated_text}'")
                        run.text = updated_text

                # Форматирование текста
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                    run.font.size = Pt(13)

                # Выравнивание
                if key.lower() == "{сегодняшняя дата 1}":
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


# Функция для замены с сохранением регистра
def replace_with_case_preservation(text, placeholder, replacement):
    def match_case(source, target):
        if source.isupper():
            return target.upper()
        elif source.islower():
            return target.lower()
        elif source.istitle():
            return target.title()
        else:
            return target

    # Поиск и замена с сохранением регистра
    start = text.lower().find(placeholder.lower())
    if start == -1:
        return text

    end = start + len(placeholder)
    original_placeholder = text[start:end]
    replacement_with_case = match_case(original_placeholder, replacement)

    return text[:start] + replacement_with_case + text[end:]

# Функция для создания PDF из DOCX
def create_pdf(docx_path, pdf_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', uni=True)
    pdf.set_font("DejaVu", size=13)

    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        pdf.multi_cell(0, 10, paragraph.text)

    pdf.output(pdf_path)

# Хендлер команды /start
@dp.message(Command("start"))
async def start(message: types.Message, state: FSMContext):
    welcome_text = (
        "🤖 **Что умеет этот бот?**\n\n"
        "Этот бот предназначен для помощи в заполнении договоров.\n\n"
        "Соберём все данные для заполнения договора шаг за шагом. Готовы начать?"
    )

    keyboard = types.ReplyKeyboardMarkup(
        keyboard=[
            [types.KeyboardButton(text="🚀 Начать заполнение договора")]
        ],
        resize_keyboard=True
    )

    await message.answer(welcome_text, reply_markup=keyboard, parse_mode="Markdown")

# Хендлер для кнопки "🚀 Начать заполнение договора"
@dp.message(lambda message: message.text == "🚀 Начать заполнение договора")
async def start_contract_filling(message: types.Message, state: FSMContext):
    await message.answer("Введите ФИО заказчика:")
    await state.set_state(ContractStates.GET_CUSTOMER_NAME)

# Обработчики ввода данных
@dp.message(ContractStates.GET_CUSTOMER_NAME)
async def get_customer_name(message: types.Message, state: FSMContext):
    await state.update_data(customer_name=message.text)
    await message.answer("Введите сумму договора (цифрами):")
    await state.set_state(ContractStates.GET_CONTRACT_AMOUNT)

@dp.message(ContractStates.GET_CONTRACT_AMOUNT)
async def get_contract_amount(message: types.Message, state: FSMContext):
    await state.update_data(contract_amount=message.text)
    await message.answer("Введите название товара в родительном падеже:")
    await state.set_state(ContractStates.GET_PRODUCT_NAME)

@dp.message(ContractStates.GET_PRODUCT_NAME)
async def get_product_name(message: types.Message, state: FSMContext):
    await state.update_data(product_name=message.text)
    await message.answer("Введите банковские реквизиты:")
    await state.set_state(ContractStates.GET_BANK_DETAILS)

# Обработчик ввода банковских реквизитов
@dp.message(ContractStates.GET_BANK_DETAILS)
async def get_bank_details(message: types.Message, state: FSMContext):
    await state.update_data(bank_details=message.text)
    data = await state.get_data()

    # Подготовка данных для заполнения
    today_date = datetime.now().strftime("%d.%m.%Y")
    future_date = (datetime.now() + timedelta(days=45)).strftime("%d.%m.%Y")

    placeholders = {
        "{сегодняшняя дата 1}": today_date,
        "{заказчик 1}": f"Индивидуальный Предприниматель {data.get('customer_name', 'Пустое значение')}",
        "{название товара в родительном падеже}": data.get('product_name', 'Пустое значение'),
        "{сегодняшняя дата}": today_date,
        "{полтора месяца вперед от сегодняшней даты}": future_date,
        "{стоимость работ цифрами}": str(data.get('contract_amount', '0')),
        "{стоимость работ прописью}": num2words(int(data.get('contract_amount', '0')), lang='ru') + " рублей 00 копеек"
    }

    logging.info(f"Передаем значения для заполнения: {placeholders}")

    doc = Document(TEMPLATE_PATH)
    replace_placeholders(doc, placeholders)

    # Сохранение DOCX
    try:
        docx_output_path = "/home/anna/syncli_doc/syncli_doc/output.docx"
        doc.save(docx_output_path)
    except Exception as e:
        logging.error(f"Ошибка при сохранении DOCX: {str(e)}")
        await message.answer(f"Ошибка при сохранении DOCX: {str(e)}")
        return

    # Создание PDF
    try:
        pdf_output_path = "/home/anna/syncli_doc/syncli_doc/output.pdf"
        create_pdf(docx_output_path, pdf_output_path)
    except Exception as e:
        logging.error(f"Ошибка при создании PDF: {str(e)}")
        await message.answer(f"Ошибка при создании PDF: {str(e)}")
        return

    # Отправка готового договора
    if os.path.exists(docx_output_path) and os.path.exists(pdf_output_path):
        try:
            await message.answer("Готовый договор создан и отправлен:")
            await message.answer_document(types.FSInputFile(docx_output_path))
            await message.answer_document(types.FSInputFile(pdf_output_path))
        except Exception as e:
            logging.error(f"Ошибка при отправке файла: {str(e)}")
            await message.answer(f"Ошибка при отправке файла: {str(e)}")
    else:
        logging.error("Ошибка: файлы не были созданы.")
        await message.answer("Ошибка: не удалось создать файлы договора.")

    await state.clear()

# Основная функция запуска бота
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
