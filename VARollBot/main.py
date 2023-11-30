from datetime import datetime

import gspread
import pytz
import telebot
from oauth2client.service_account import ServiceAccountCredentials
from telebot import types

# utc_now = datetime.utcnow()
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Read about loggers, ask gpt, or me to provide u some links

# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

 
# Замените на свой токен бота и данные для доступа к гугл-таблице


# Инициализация бота и доступа к гугл-таблице
bot = telebot.TeleBot(TOKEN)
creds = ServiceAccountCredentials.from_json_keyfile_name("exeljs.json")
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_KEY)
worksheet = spreadsheet.get_worksheet(0)  # Предполагается, что данные будут записываться в первый лист таблицы

# Список доступных типов
available_buttons = ["опоздание//задержка на работе//отсутствие", "Прервался звонок"]

# Словарь для временного хранения данных пользователя
user_data = {}


msk_tz = pytz.timezone("Europe/Moscow")

# id Telegram пользователя, которого нужно уведомлять



# Обработчик команды /start
@bot.message_handler(commands=["start"])
def start(message):
    bot.send_message(message.chat.id, "Привет! Я бот для фиксации информации в гугл-таблицу. " "Отправь мне /record для начала записи.")


# Обработчик команды /record
@bot.message_handler(commands=["record"])
def record(message):
    markup = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    for record_type in available_buttons:
        markup.add(types.KeyboardButton(record_type))
    bot.send_message(message.chat.id, "Выберите тип:", reply_markup=markup)
    bot.register_next_step_handler(message, process_type_step)


# Обработчик выбора типа
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


def get_or_create_monthly_sheet(current_datetime):
    month_year = current_datetime.strftime("%B %Y")
    try:
        return spreadsheet.worksheet(month_year)
    except gspread.exceptions.WorksheetNotFound:
        # Если листа для текущего месяца нет, создаем новый
        new_sheet = spreadsheet.add_worksheet(month_year, rows="100", cols="20")
        new_sheet.append_row(["Дата", "Пользователь", "Тип", "Комментарий"])  # Заголовки столбцов
        return new_sheet

# Функция для добавления записи в лист для текущего месяца
def add_record_to_monthly_sheet(monthly_sheet, current_datetime, user_info, user_type, user_comment):
    date = current_datetime.strftime("%Y-%m-%d %H:%M:%S")
    data = [date, user_info, user_type, user_comment]
    monthly_sheet.append_row(data)


# Обработчик ввода комментария
def process_comment_step(message):
    try:
        # Получаем введенный комментарий
        user_comment = message.text.strip()

        # Получаем информацию о пользователе
        user_info = f"@{message.from_user.username}"

        # Получаем текущую дату и время в часовом поясе Москвы
        current_datetime = datetime.now(msk_tz)

        # Получаем или создаем лист для текущего месяца
        monthly_sheet = get_or_create_monthly_sheet(current_datetime)

        # Если текущий лист пустой, добавляем запись с датой текущего дня

        # Добавляем запись в лист для текущего месяца
        add_record_to_monthly_sheet(monthly_sheet, current_datetime, user_info, user_data["type"], user_comment)

        # Отправляем уведомление пользователю
        notification_message = (
            f"Информация успешно записана в гугл-таблицу! "
            f"Тип: {user_data['type']}, Дата: {current_datetime.strftime('%Y-%m-%d %H:%M:%S')}, Комментарий: {user_comment}"
        )
        bot.send_message(message.chat.id, notification_message)

        # Отправляем уведомление администратору (в данном случае, пользователю @mobessona)
        message_for_admin = f"Пользователь {user_info} записал информацию в гугл-таблицу.\n{notification_message}"
        bot.send_message(admin_username, message_for_admin)

    except gspread.exceptions.APIError as e:
        print(f"Google Sheets API Error: {e}")
        error_message = "Произошла ошибка при записи данных в гугл-таблицу. Пожалуйста, попробуйте еще раз или обратитесь к администратору."
        bot.send_message(message.chat.id, error_message)

    except Exception as e:
        print(f"Error: {e}")
        error_message = "Произошла ошибка. Пожалуйста, попробуйте еще раз или обратитесь к администратору."
        bot.send_message(message.chat.id, error_message)


# Запуск бота
if __name__ == "__main__":
    bot.polling(none_stop=True)
bot.infinity_polling()
