import telebot
import openpyxl
import datetime
import os

# Get the value of the environment variable
TOKEN = os.getenv("TOKEN")

# Load the Excel file
workbook = openpyxl.load_workbook('03.09.xlsx')

# Select the sheet with the schedule
sheet = workbook['1']  # Replace with the name of your sheet

# Get the current day of the week
now = datetime.datetime.now()
day_name = now.strftime('%A')
# Создаем словарь с соответствиями для русских названий дней недели
day_names_dict = {'Monday': 'Понедельник', 'Tuesday': 'Вторник', 'Wednesday': 'Среда', 'Thursday': 'Четверг', 'Friday': 'Пятница', 'Saturday': 'Суббота', 'Sunday': 'Воскресенье'}
# Заменяем название дня недели на русское название
day_name_ru = day_names_dict.get(day_name, day_name)

# Порядковый номер предмета к расписанию звонков 
number_to_time_dict_mon = {'1' : '8.00 - 9.20', '2' : '9.35 - 10.55', '3' : '11.10 - 12-30',
        '4' : '13.40 - 15.00', '5' : '15.15 - 16.35', '6' : '16.50 - 18.10'}
number_to_time_dict_wed = {'1' : '8.00 - 9.30', '2' : '9.45 - 11.15', '3' : '11.30 - 13.00',
        '4' : '13.15 - 14.45', '5' : '15.00 - 16.30', '6' : '16.40 - 18.10'}
number_to_time_dict_sat = {'1' : '8.00 - 9.20', '2' : '9.30 - 10.50', '3' : '11.00 - 12.20',
        '4' : '12.40 - 14.00', '5' : '14.10 - 15.30', '6' : '15.40 - 17.00'}

# Setup the bot token
bot = telebot.TeleBot(TOKEN)

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, "Бот с расписанием, чтобы начать напишите  /start ")
   # if not str(message.chat.id) in joinedUsers:
    #    joinedFile = open('ids.txt', 'a')
     #   joinedFile.write(str(message.chat.id) + "\n")
      #  joinedUsers.add(message.chat.id)

#@bot.message_handler(commands=['start'])
#def select_department(message):
#    bot.send_message

# Define a function to handle incoming messages
@bot.message_handler(func=lambda message: True)
def handle_message(message):
    group_number = message.text
    group_row = None
    for row in sheet.iter_rows(min_row=6, max_row=sheet.max_row, min_col=1, max_col=1):
        if row[0].value == f'Группа - {group_number}':
            group_row = row[0].row
            break
        elif row[0].value == f'Группа -{group_number}':
            group_row = row[0].row
            break

    current_day_col = None
    for col in sheet.iter_cols(min_row=7, max_row=7, min_col=2, max_col=sheet.max_column):
        if col[0].value == day_name_ru:
            current_day_col = col[0].column
            break

    if group_row is not None and current_day_col is not None:
        schedule = f'Расписание занятий для группы - {group_number} на {day_name_ru}: \n'
        item_number = 0
        for row in sheet.iter_rows(min_row=group_row+3, max_row=group_row+8, min_col=current_day_col, max_col=current_day_col+2):
            lesson_number = row[0].value
            subject_name = row[1].value
            item_number += 1
            if day_name_ru in ['Понедельник', 'Вторник']:
                number_to_time_dict = number_to_time_dict_mon
            elif day_name_ru in ['Среда', 'Четверг', 'Пятница']:
                number_to_time_dict = number_to_time_dict_wed
            else:
                number_to_time_dict = number_to_time_dict_sat
            number_to_time = number_to_time_dict.get(str(item_number), str(item_number))
            if lesson_number is not None:
                schedule += f'{item_number}. {number_to_time}. {lesson_number}. {subject_name}\n'
        bot.send_message(message.chat.id, schedule)
    else:
        if group_row is None:
            bot.send_message(message.chat.id, f'Группа - {group_number} not found in the schedule.')
        if current_day_col is None:
            bot.send_message(message.chat.id, f'No class schedule found for {day_name_ru}.')

# Start the bot
bot.polling()

