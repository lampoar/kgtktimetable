import os
import telebot
import openpyxl

# Load the token from the environment variable
TOKEN = os.getenv('TOKEN')

# Create the bot
bot = telebot.TeleBot(TOKEN)

# Define the available departments
departments = {
    'Педагогическое': 'pedagog.xlsx',
    'Строительное': 'building.xlsx',
    'Технологическое': 'technology.xlsx'
}

@bot.message_handler(commands=['start'])
def handle_start(message):
    # Send a message with the department selection buttons
    keyboard = telebot.types.InlineKeyboardMarkup(row_width=2)
    for department in departments:
        callback_data = f"department_selection_{department}"
        button = telebot.types.InlineKeyboardButton(department, callback_data=callback_data)
        keyboard.add(button)
    bot.send_message(message.chat.id, 'Выберите ваше отделение:', reply_markup=keyboard)

@bot.callback_query_handler(func=lambda call: call.data.startswith('department_selection_'))
def handle_department_selection(call):
    # Load the workbook for the selected department
    department = call.data[len('department_selection_'):]
    workbook_name = departments[department]
    workbook = openpyxl.load_workbook(workbook_name)

    # Get the list of group names from the workbook
    sheet = workbook.active
    group_names = [cell.value for cell in sheet['A'] if cell.value is not None and str(cell.value.startswith('Группа - '))]

    # Send a message with the group selection buttons
    keyboard = telebot.types.InlineKeyboardMarkup(row_width=2)
    for group_name in group_names:
        group_number = group_name[len('Группа - '):]
        callback_data = f"group_selection_{group_number}"
        button = telebot.types.InlineKeyboardButton(group_name, callback_data=callback_data)
        keyboard.add(button)
    bot.send_message(call.message.chat.id, 'Выберите вашу группу:', reply_markup=keyboard)

@bot.callback_query_handler(func=lambda call: call.data.startswith('group_selection_'))
def handle_group_selection(call):
    # Get the selected group number
    group_number = call.data[len('group_selection_'):]

    # TODO: Load the class schedule for the selected group and send it to the user

# Start the bot
bot.polling()

