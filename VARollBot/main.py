from datetime import datetime

import gspread
import pytz
import telebot
from oauth2client.service_account import ServiceAccountCredentials
from telebot import types

SPREADSHEET_KEY = ***REMOVED***
INCIDENTS_SHEET_NAME = "incidents"
STAFF_SHEET_NAME = "staff"
ADMINS_IDS = ***REMOVED***


bot = telebot.TeleBot(TOKEN)
creds = ServiceAccountCredentials.from_json_keyfile_name("exeljs.json")
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_KEY)
worksheet = spreadsheet.get_worksheet(0)


available_buttons = ["опоздание//задержка на работе//отсутствие", "Прервался звонок", "Cторонний функционал"]
user_data = {}
current_date = None
msk_tz = pytz.timezone("Europe/Moscow")


@bot.message_handler(commands=["start"])
def start(message):
    bot.send_message(message.chat.id, "Привет! Я бот для фиксации информации в гугл-таблицу. " "Отправь мне /record для начала записи.")


@bot.message_handler(commands=["record"])
def record(message):
    markup = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    for record_type in available_buttons:
        markup.add(types.KeyboardButton(record_type))
    bot.send_message(message.chat.id, "Выберите тип:", reply_markup=markup)
    bot.register_next_step_handler(message, process_type_step)


def process_type_step(message):
    try:
        # Проверяем, что тип входит в список доступных
        if message.text in available_buttons:
            # Сохраняем выбранный тип
            user_type = message.text
            user_data["type"] = user_type

            # Отправляем запрос на ввод комментария
            bot.send_message(message.chat.id, "Введите комментарий:", reply_markup=types.ReplyKeyboardRemove())
            bot.register_next_step_handler(message, process_comment_step)
        else:
            bot.send_message(message.chat.id, "Выберите тип, используя кнопки на клавиатуре.")
    except Exception as e:
        print(f"Error: {e}")
        bot.send_message(message.chat.id, "Произошла ошибка. Пожалуйста, попробуйте еще раз.")


def get_or_create_incidents_sheet():
    try:
        return spreadsheet.worksheet(INCIDENTS_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        # Если лист "incidents" не существует, создаем новый
        new_sheet = spreadsheet.add_worksheet(INCIDENTS_SHEET_NAME, rows="100", cols="4")
        new_sheet.append_row(["Дата", "Пользователь", "Тип", "Комментарий"])
        return new_sheet


def get_admins_ids() -> list[int]:
    try:
        staff_sheet = spreadsheet.worksheet(STAFF_SHEET_NAME)
        admin_username_str = int(staff_sheet.acell("B1").value)
        seckond_admin_username_str = int(staff_sheet.acell("B2").value)

        return [admin_username_str, seckond_admin_username_str]
    except Exception as e:
        print(f"Ошибка при получении имени администратора: {e}")
        return ADMINS_IDS


# Обработчик ввода комментария
def process_comment_step(message):
    try:
        # Получаем введенный комментарий
        user_comment = message.text.strip()

        # Получаем информацию о пользователе
        user_info = f"@{message.from_user.username}"

        # Проверяем и обнуляем дату текущего дня при необходимости
        check_and_reset_current_date()

        # Получаем текущую дату и время в часовом поясе Москвы
        current_datetime = datetime.now(msk_tz)

        # Получаем лист сотрудников
        staff_sheet = spreadsheet.worksheet(STAFF_SHEET_NAME)

        cell = None
        for row in staff_sheet.findall(user_info, in_column=2):
            if row:
                cell = row
                break

        if cell:
            # Если пользователь найден, заменяем user_info значением из столбца A
            user_info = staff_sheet.cell(cell.row, 1).value

        # Добавьте следующие строки для получения или создания листа инцидентов
        incident_sheet = get_or_create_incidents_sheet()
        incident_data = incident_sheet.get_all_values()

        # Получаем последнюю дату из таблицы, если есть данные
        last_row_date = incident_data[-1][0].split()[0] if incident_data else None

        if last_row_date == current_date:
            # Если последняя дата совпадает с датой инцидента, добавляем новую строку с данными инцидента.
            incident_sheet.append_row([current_datetime.strftime("%Y-%m-%d %H:%M:%S"), user_info, user_data["type"], user_comment])
        else:
            # Если таблица пуста или последняя дата не совпадает с датой инцидента,
            # добавляем новую строку с текущей датой и объединяем четыре ячейки.
            incident_sheet.append_row([current_date])

            # Получаем номер добавленной строки
            last_row_number = len(incident_data) + 1

            # Объединяем четыре ячейки для строки с текущей датой.
            cell_range_date = f"A{last_row_number}:D{last_row_number}"
            incident_sheet.merge_cells(cell_range_date, merge_type="MERGE_ALL")

            # Обновляем форматирование для строки с текущей датой (центрирование).
            cell_format_date = {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}
            incident_sheet.format(cell_range_date, cell_format_date)

            # Добавляем новую строку с данными инцидента.
            incident_sheet.append_row([current_datetime.strftime("%Y-%m-%d %H:%M:%S"), user_info, user_data["type"], user_comment])

        # Отправляем уведомление пользователю
        notification_message = (
            f"Информация успешно записана в гугл-таблицу! "
            f"Тип: {user_data['type']}, Дата: {current_datetime.strftime('%Y-%m-%d %H:%M:%S')}, Комментарий: {user_comment}"
        )
        bot.send_message(message.chat.id, notification_message)

        # Отправляем уведомление администратору
        admins_ids = get_admins_ids()
        message_for_admin = f"Пользователь {user_info} записал информацию в гугл-таблицу.\n{notification_message}"
        for admin_id in admins_ids:
            bot.send_message(admin_id, message_for_admin)

    except Exception as e:
        print(f"Error: {e}")
        error_message = "Произошла ошибка. Пожалуйста, попробуйте еще раз или обратитесь к администратору."
        bot.send_message(message.chat.id, error_message)


def check_and_reset_current_date():
    global current_date
    today = datetime.now(msk_tz).strftime("%Y-%m-%d")
    if current_date != today:
        current_date = today


if __name__ == "__main__":
    bot.polling(none_stop=True)
