import telebot
import openpyxl
import datetime
import sqlite3

conn = sqlite3.connect('data.db')
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS users
             (chat_id INTEGER PRIMARY KEY, group_number INTEGER)''')

conn.commit()
conn.close()
# Load the Excel file
workbook = openpyxl.load_workbook('ped.xlsx')

# Select the sheet with the schedule
sheet = workbook['1']  # Replace with the name of your sheet

now = None
# Get the current day of the week
now = datetime.datetime.now()

# Get the day of the week for today and tomorrow
day_names_dict = {'Monday': 'Понедельник', 'Tuesday': 'Вторник', 'Wednesday': 'Среда', 'Thursday': 'Четверг',
                  'Friday': 'Пятница', 'Saturday': 'Суббота', 'Sunday': 'Воскресенье'}

today_day_name = now.strftime('%A')
today_day_name_ru = day_names_dict.get(today_day_name, today_day_name)

tomorrow = now + datetime.timedelta(days=1)
tomorrow_day_name = tomorrow.strftime('%A')
tomorrow_day_name_ru = day_names_dict.get(tomorrow_day_name, tomorrow_day_name)

# Порядковый номер предмета к расписанию звонков
# На Понедельник. Вторник
number_to_time_dict_mon = {'1': '8.00 - 9.20', '2': '9.35 - 10.55', '3': '11.10 - 12-30',
                           '4': '13.40 - 15.00', '5': '15.15 - 16.35', '6': '16.50 - 18.10'}
number_to_time_dict_wed = {'1': '8.00 - 9.30', '2': '9.45 - 11.15', '3': '11.30 - 13.00',
                           '4': '13.15 - 14.45', '5': '15.00 - 16.30', '6': '16.40 - 18.10'}
number_to_time_dict_sat = {'1': '8.00 - 9.20', '2': '9.30 - 10.50', '3': '11.00 - 12.20',
                           '4': '12.40 - 14.00', '5': '14.10 - 15.30', '6': '15.40 - 17.00'}

# Create the Telegram bot object
bot = telebot.TeleBot('2022385697:AAGiFoszmY_6jzPm9DLhf9P1gjz9PtNpKuw')


# Define the start command
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message,
                 'Привет! Я помогу тебе узнать расписание занятий для твоей группы. Введи номер группы, например "123", Для того чтобы сменить группу введи /chgroupe *номер группы*.')


# Define the message handler
@bot.message_handler(func=lambda message: True )
def echo_all(message):
    # Check if the user has already entered their group number
    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    c.execute('SELECT * FROM users WHERE chat_id = ?', (message.chat.id,))
    result = c.fetchone()

    if result is None:
        # The user has not entered their group number yet
        group_number = message.text
        c.execute('INSERT INTO users VALUES (?, ?)', (message.chat.id, group_number))
        conn.commit()
    else:
        group_number = result[1]

    process_day_choice(message, group_number)


# Define the day choice handler
def process_day_choice(message, group_number):
  #  if
    # Ask the user if they want to see today's schedule, tomorrow's schedule or the whole week's schedule
    markup = telebot.types.ReplyKeyboardMarkup(row_width=1)
    itembtn1 = telebot.types.KeyboardButton('Сегодня')
    itembtn2 = telebot.types.KeyboardButton('Завтра')
    itembtn3 = telebot.types.KeyboardButton('Неделя')
    markup.add(itembtn1, itembtn2, itembtn3)

    msg = bot.reply_to(message,
                       f'Хотите посмотреть расписание на {today_day_name_ru}, на {tomorrow_day_name_ru} или на всю неделю?',
                       reply_markup=markup)

    day_choice = message.text.lower()

    if day_choice == 'сегодня':
        day_name_ru = today_day_name_ru
        days = [day_name_ru]
        process_schedule(message, group_number, days)
    elif day_choice == 'завтра':
        day_name_ru = tomorrow_day_name_ru
        days = [day_name_ru]
        process_schedule(message, group_number, days)
    elif day_choice == 'неделя':
        day_name_ru = list(day_names_dict.values())
        days = day_name_ru
        process_schedule(message, group_number, days)
 #   else:
 #       bot.send_message(message.chat.id, 'Некорректный ввод. Попробуйте еще раз.')


def process_schedule(message, group_number, days):
    # Find the row with the group number
    group_row = None
    for row in sheet.iter_rows(min_row=6, max_row=sheet.max_row, min_col=1, max_col=1):
        if row[0].value == f'Группа - {group_number}':
            group_row = row[0].row
            break
        elif row[0].value == f'Группа -{group_number}':
            group_row = row[0].row
            break

    # Send the class schedule for the chosen day(s) of the week for the given group
    if group_row is not None:
        schedule = f'Расписание занятий для группы - {group_number}:\n'
        for day_name_ru in days:
            # Find the column with the chosen day of the week
            day_col = None
            for col in sheet.iter_cols(min_row=7, max_row=7, min_col=2, max_col=sheet.max_column):
                if col[0].value == day_name_ru:
                    day_col = col[0].column
                    break
            if day_col is not None:
                schedule += f'\n{day_name_ru}:\n'
                item_number = 0
                for row in sheet.iter_rows(min_row=group_row+3, max_row=group_row+8, min_col=day_col, max_col=day_col+2):
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
                        schedule += f'{number_to_time}. {lesson_number}. {subject_name}\n'
            else:
                schedule += f'Расписание занятий на {day_name_ru} не найдено.\n'
        bot.send_message(message.chat.id, schedule)
    else:
        bot.send_message(message.chat.id, f'Группа - {group_number} не найдена в расписании.')


@bot.message_handler(commands=['chgroupe'])
def change_group_number(message):
    group_number = message.text.split('/chgroupe ')[1]

    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    c.execute('SELECT * FROM users WHERE chat_id = ?', (message.chat.id,))
    result = c.fetchone()

    if result is None:
        # The user has not entered their group number yet
        c.execute('INSERT INTO users VALUES (?, ?)', (message.chat.id, group_number))
    else:
        c.execute('UPDATE users SET group_number = ? WHERE chat_id = ?', (group_number, message.chat.id))

    conn.commit()
    conn.close()

    bot.reply_to(message, f'Номер группы успешно изменен на {group_number}.')


# Start the bot
bot.polling()