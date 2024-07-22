from datetime import datetime

import gspread
import pytz
import telebot
from oauth2client.service_account import ServiceAccountCredentials
from telebot import types

TOKEN = 
SPREADSHEET_KEY = 
INCIDENTS_SHEET_NAME = "incidents"
STAFF_SHEET_NAME = "staff"
ADMINS_IDS = 

bot = telebot.TeleBot(TOKEN)
creds = ServiceAccountCredentials.from_json_keyfile_name("exeljs.json")
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_KEY)
worksheet = spreadsheet.get_worksheet(0)

available_buttons = ["late/delay at work/absence", "Call disconnected", "Third-party functionality"]
user_data = {}
current_date = None
msk_tz = pytz.timezone("Europe/Moscow")

@bot.message_handler(commands=["start"])
def start(message):
    bot.send_message(message.chat.id, "Hello! I'm a bot for logging information to a Google Sheet. Send me /record to start recording.")

@bot.message_handler(commands=["record"])
def record(message):
    markup = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    for record_type in available_buttons:
        markup.add(types.KeyboardButton(record_type))
    bot.send_message(message.chat.id, "Select type:", reply_markup=markup)
    bot.register_next_step_handler(message, process_type_step)

def process_type_step(message):
    try:
        if message.text in available_buttons:
            user_type = message.text
            user_data["type"] = user_type
            bot.send_message(message.chat.id, "Enter comment:", reply_markup=types.ReplyKeyboardRemove())
            bot.register_next_step_handler(message, process_comment_step)
        else:
            bot.send_message(message.chat.id, "Select type using the buttons on the keyboard.")
    except Exception as e:
        print(f"Error: {e}")
        bot.send_message(message.chat.id, "An error occurred. Please try again.")

def get_or_create_incidents_sheet():
    try:
        return spreadsheet.worksheet(INCIDENTS_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        new_sheet = spreadsheet.add_worksheet(INCIDENTS_SHEET_NAME, rows="100", cols="4")
        new_sheet.append_row(["Date", "User", "Type", "Comment"])
        return new_sheet

def get_admins_ids() -> list[int]:
    try:
        staff_sheet = spreadsheet.worksheet(STAFF_SHEET_NAME)
        admin_username_str = int(staff_sheet.acell("B1").value)
        seckond_admin_username_str = int(staff_sheet.acell("B2").value)
        return [admin_username_str, seckond_admin_username_str]
    except Exception as e:
        print(f"Error fetching admin name: {e}")
        return ADMINS_IDS

def process_comment_step(message):
    try:
        user_comment = message.text.strip()
        user_info = f"@{message.from_user.username}"
        check_and_reset_current_date()
        current_datetime = datetime.now(msk_tz)
        staff_sheet = spreadsheet.worksheet(STAFF_SHEET_NAME)
        cell = None
        for row in staff_sheet.findall(user_info, in_column=2):
            if row:
                cell = row
                break
        if cell:
            user_info = staff_sheet.cell(cell.row, 1).value
        incident_sheet = get_or_create_incidents_sheet()
        incident_data = incident_sheet.get_all_values()
        last_row_date = incident_data[-1][0].split()[0] if incident_data else None
        if last_row_date == current_date:
            incident_sheet.append_row([current_datetime.strftime("%Y-%m-%d %H:%M:%S"), user_info, user_data["type"], user_comment])
        else:
            incident_sheet.append_row([current_date])
            last_row_number = len(incident_data) + 1
            cell_range_date = f"A{last_row_number}:D{last_row_number}"
            incident_sheet.merge_cells(cell_range_date, merge_type="MERGE_ALL")
            cell_format_date = {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}
            incident_sheet.format(cell_range_date, cell_format_date)
            incident_sheet.append_row([current_datetime.strftime("%Y-%m-%d %H:%M:%S"), user_info, user_data["type"], user_comment])
        notification_message = (f"Information successfully recorded in Google Sheet! Type: {user_data['type']}, Date: {current_datetime.strftime('%Y-%m-%d %H:%M:%S')}, Comment: {user_comment}")
        bot.send_message(message.chat.id, notification_message)
        admins_ids = get_admins_ids()
        message_for_admin = f"User {user_info} logged information in Google Sheet.\n{notification_message}"
        for admin_id in admins_ids:
            bot.send_message(admin_id, message_for_admin)
    except Exception as e:
        print(f"Error: {e}")
        error_message = "An error occurred. Please try again or contact an administrator."
        bot.send_message(message.chat.id, error_message)

def check_and_reset_current_date():
    global current_date
    today = datetime.now(msk_tz).strftime("%Y-%m-%d")
    if current_date != today:
        current_date = today

if __name__ == "__main__":
    bot.polling(none_stop=True)
