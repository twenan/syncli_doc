import asyncio
import os
import datetime
import logging
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt  # Добавляем импорт для размера шрифта
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
    # Замена текста в параграфах
    for paragraph in doc.paragraphs:
        for key, value in placeholders.items():
            if key in paragraph.text:
                inline = paragraph.runs
                for i in range(len(inline)):
                    if key in inline[i].text:
                        inline[i].text = inline[i].text.replace(key, value)

                # Устанавливаем шрифт Times New Roman
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                    run.font.size = Pt(13)

                # Выравнивание
                if key == "{Сегодняшняя дата 1}":
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif key in [
                    "{Заказчик 1}",
                    "{Название товара в родительном падеже}",
                    "{Сегодняшняя дата}",
                    "{Полтора месяца вперед от сегодняшней даты}",
                    "{Сумма работ цифрами}",
                    "{Сумма работ прописью}",
                ]:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Замена текста в таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in placeholders.items():
                    if key in cell.text:
                        new_text = cell.text.replace(key, value)
                        cell.text = new_text

                        # Устанавливаем шрифт и форматирование для текста в таблицах
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.name = 'Times New Roman'
                                run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                                run.font.size = Pt(13)

                        # Применяем выравнивание для каждой ячейки
                        for paragraph in cell.paragraphs:
                            if key == "{Сегодняшняя дата 1}":
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            elif key in [
                                "{Заказчик 1}",
                                "{Название товара в родительном падеже}",
                                "{Сегодняшняя дата}",
                                "{Полтора месяца вперед от сегодняшней даты}",
                                "{Сумма работ цифрами}",
                                "{Сумма работ прописью}",
                            ]:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# Функция для создания PDF из DOCX
def create_pdf(docx_path, pdf_path):
    pdf = FPDF()
    pdf.add_page()

    # Подключаем шрифт с поддержкой кириллицы
    pdf.add_font('DejaVu', '', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', uni=True)
    pdf.set_font("DejaVu", size=13)

    # Читаем текст из DOCX
    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        pdf.multi_cell(0, 10, paragraph.text)

    # Сохраняем PDF
    pdf.output(pdf_path)

# Хендлер команды /start
@dp.message(Command("start"))
async def start(message: types.Message, state: FSMContext):
    welcome_text = (
        "🤖 **Что умеет этот бот?**\n\n"
        "Этот бот предназначен для помощи в заполнении договоров.\n\n"
        "Соберём все данные для заполнения договора шаг за шагом. Готовы начать?"
    )

    # Создание клавиатуры
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

# Обработчик ввода банковских реквизитов и создание договора
@dp.message(ContractStates.GET_BANK_DETAILS)
async def get_bank_details(message: types.Message, state: FSMContext):
    await state.update_data(bank_details=message.text)
    data = await state.get_data()

    # Текущая дата
    today_date = datetime.now().strftime("%d.%m.%Y")
    # Дата через 45 дней
    future_date = (datetime.now() + timedelta(days=45)).strftime("%d.%m.%Y")

    # Заполнение шаблона
    doc = Document(TEMPLATE_PATH)
    placeholders = {
        "{Сегодняшняя дата 1}": today_date,
        "{Заказчик 1}": f"Индивидуальный Предприниматель {data['customer_name']}",
        "{Название товара в родительном падеже}": data['product_name'],
        "{Сегодняшняя дата}": today_date,
        "{Полтора месяца вперед от сегодняшней даты}": future_date,
        "{Сумма работ цифрами}": data['contract_amount'],
        "{Сумма работ прописью}": num2words(int(data['contract_amount']), lang='ru') + " рублей 00 копеек"
    }

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

    # Проверка существования файлов
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
