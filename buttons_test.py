import telebot
import openpyxl
import datetime
import os
import sqlite3

TOKEN = os.getenv("TOKEN")
bot = telebot.TeleBot(TOKEN)


@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message,
                 'Привет! Я помогу тебе узнать расписание занятий для твоей группы. Для смены группы введи /chgroup '
                 '№Группы, например /chgroup 1234. Больше информации в /about. Выберите отделение:',
                 reply_markup=get_department_keyboard(message))
def get_department_keyboard(message):
    markup = telebot.types.ReplyKeyboardMarkup(row_width=1)
    itembtn1 = telebot.types.KeyboardButton('Педагогическое')
    itembtn2 = telebot.types.KeyboardButton('Технологическое')
    itembtn3 = telebot.types.KeyboardButton('Строительное')
    markup.add(itembtn1, itembtn2, itembtn3)
    return markup
