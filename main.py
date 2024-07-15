from telebot import TeleBot, types
import requests
import json
from openpyxl import load_workbook
import datetime
import schedule
import threading
import time

bot = TeleBot('7420715170:AAFQhpQFvdYi4fTdZULCSKilYH22vsd6dFo')
API = '008c3f86611c8333d260684d8c4fe8c4'
city = 'Kerch'
file = load_workbook(filename='iit_23_b_o28.05.xlsx')
day = datetime.date.today()
today = day.strftime("%d.%m.%Y")

@bot.message_handler(commands=['start'])
def start(message):
    murkup = types.ReplyKeyboardMarkup()
    buttom1 = types.KeyboardButton('Сайт расписания')
    buttom2 = types.KeyboardButton('Погода')
    buttom3 = types.KeyboardButton('Расписание на сегодня')
    murkup.row(buttom1)
    murkup.row(buttom2)
    murkup.row(buttom3)
    bot.send_message(message.chat.id, f'Привет {message.from_user.first_name} {message.from_user.last_name}', reply_markup=murkup)

@bot.message_handler(content_types=['text'])
def proverka(message):
    if message.text == "Сайт расписания":
        bot.send_message(message.chat.id, f'https://timetable.sevsu.ru/auth/keycloak/redirect')
    elif message.text == 'Погода':
        weather(message)
    elif message.text == 'Расписание на сегодня':
        raspisanie(message)
    elif message.text == 'Расписание на завтра':
        raspisanie(message)

def raspisanie(message=None):
    for a in range(23, 42):
        list = file['уч.н. ' + str(a)]
        for i in range(7, 56, 8):
            if list['B' + str(i)].value == today:
                for n in range(i, i + 8):
                    noneF = 'F'
                    noneG = 'G'
                    if list['F' + str(n)].value is None:
                        noneF = "I"
                        noneG = "J"
                    if list['E' + str(n)].value is None:
                        i += 8
                    else:
                        message_text = (f"{list['C' + str(n)].value} пара, время: {list['D' + str(n)].value}\n\n"
                                        f"{list['E' + str(n)].value}\nТип: {list[noneF + str(n)].value}\nАудитория {list[noneG + str(n)].value}\n\n")
                        if message:
                            bot.send_message(message.chat.id, message_text)
                        else:
                            # Замените YOUR_CHAT_ID на ID чата, куда должно отправляться сообщение
                            bot.send_message('YOUR_CHAT_ID', message_text)

def weather(message):
    res = requests.get(f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid={API}&units=metric')
    data = json.loads(res.text)
    temp = data["main"]["temp"]
    Weat = data["weather"][0]["main"]
    if Weat == "Clouds":
        Weat = 'Облачно'
    elif Weat == "Clear":
        Weat = "Солнечно"
    elif Weat == "Rainy":
        Weat = "Дождь"
    wind_speed = data["wind"]["speed"]
    bot.send_message(message.chat.id, f'Сейчас {Weat}, {temp} градусов\nВетер {wind_speed} метров/с\n')

# Выдаёт всю инфу про чела и чат
@bot.message_handler(commands=['about_user'])
def main(message):
    bot.send_message(message.chat.id, message)

def schedule_task():
    schedule.every().day.at("22:17").do(raspisanie)
    while True:
        schedule.run_pending()
        time.sleep(1)

# Запускаем планировщик в отдельном потоке
schedule_thread = threading.Thread(target=schedule_task)
schedule_thread.start()

# Это чтобы бот не отключался, а работал. По сути вечный цикл
bot.infinity_polling()
