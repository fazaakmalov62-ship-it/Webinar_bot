import telebot
from telebot import types
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

# === Настройки ===
BOT_TOKEN = "7833833174:AAHOVAge-5gsRPWbYj85Wd4WJRevYMqU5wg"
ADMIN_ID = 1309971729  # замените на ваш Telegram ID
FILE_NAME = "webinar_registrations.xlsx"

bot = telebot.TeleBot(BOT_TOKEN)

# === Создаём файл Excel, если его нет ===
if not os.path.exists(FILE_NAME):
    wb = Workbook()
    ws = wb.active
    ws.append(["Никнейм", "Дата регистрации", "ID TG", "Имя", "Статус"])
    wb.save(FILE_NAME)

# === Стартовое сообщение ===
@bot.message_handler(commands=['start'])
def start(message):
    user_id = message.from_user.id
    wb = load_workbook(FILE_NAME)
    ws = wb.active

    # Проверяем, есть ли пользователь
    found = False
    for i in range(2, ws.max_row + 1):
        if ws.cell(row=i, column=3).value == user_id:
            found = True
            name = ws.cell(row=i, column=4).value or "—"
            status = ws.cell(row=i, column=5).value or ""
            break

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add(types.KeyboardButton("Записаться на вебинар"))
    markup.add(types.KeyboardButton("Обновить данные"))
    markup.add(types.KeyboardButton("Отказаться от вебинара"))

    if found:
        bot.send_message(message.chat.id,
                         f"👋 Вы уже зарегистрированы!\nИмя: {name}\nСтатус: {status}",
                         reply_markup=markup)
    else:
        bot.send_message(message.chat.id,
                         "👋 Привет! Нажмите кнопку ниже, чтобы записаться на вебинар:",
                         reply_markup=markup)

# === Нажатие «Записаться на вебинар» ===
@bot.message_handler(func=lambda msg: msg.text == "Записаться на вебинар")
def register_step1(message):
    user_id = message.from_user.id
    username = message.from_user.username or "—"
    reg_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    wb = load_workbook(FILE_NAME)
    ws = wb.active

    # Проверяем, есть ли пользователь
    found = False
    for i in range(2, ws.max_row + 1):
        if ws.cell(row=i, column=3).value == user_id:
            # Обновляем никнейм и дату регистрации
            ws.cell(row=i, column=1, value=username)
            ws.cell(row=i, column=2, value=reg_time)
            found = True
            break

    if not found:
        ws.append([username, reg_time, user_id, "", ""])
        bot.send_message(ADMIN_ID, f"Новый участник зарегистрировался: @{username} ({user_id})")

    wb.save(FILE_NAME)

    bot.send_message(message.chat.id, "✍️ Пожалуйста, введите ваше имя:")
    bot.register_next_step_handler(message, register_step2)

# === Получаем имя ===
def register_step2(message):
    name = message.text
    user_id = message.from_user.id

    wb = load_workbook(FILE_NAME)
    ws = wb.active
    for i in range(2, ws.max_row + 1):
        if ws.cell(row=i, column=3).value == user_id:
            ws.cell(row=i, column=4, value=name)
            ws.cell(row=i, column=5, value="")  # статус пустой
            break
    wb.save(FILE_NAME)

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add(types.KeyboardButton("Отказаться от вебинара"))
    markup.add(types.KeyboardButton("Обновить данные"))
    bot.send_message(message.chat.id,
                     "✅ Спасибо за регистрацию! Ваши данные сохранены.",
                     reply_markup=markup)

# === Отказ от вебинара ===
@bot.message_handler(func=lambda msg: msg.text == "Отказаться от вебинара")
def cancel_registration(message):
    user_id = message.from_user.id

    wb = load_workbook(FILE_NAME)
    ws = wb.active
    for i in range(2, ws.max_row + 1):
        if ws.cell(row=i, column=3).value == user_id:
            ws.cell(i, 5, "Отказался")
            break
    wb.save(FILE_NAME)
    bot.send_message(message.chat.id, "❌ Вы отказались от вебинара. Данные обновлены.")

# === Обновление данных ===
@bot.message_handler(func=lambda msg: msg.text == "Обновить данные")
def update_data(message):
    # Просто начинаем регистрацию заново для обновления данных
    register_step1(message)

# === Команда админа для рассылки ===
@bot.message_handler(commands=['broadcast'])
def broadcast(message):
    if message.from_user.id != ADMIN_ID:
        bot.send_message(message.chat.id, "❌ У вас нет доступа.")
        return
    msg = bot.send_message(message.chat.id, "Введите текст рассылки:")
    bot.register_next_step_handler(msg, send_broadcast)

def send_broadcast(message):
    text = message.text
    wb = load_workbook(FILE_NAME)
    ws = wb.active
    sent = 0
    for i in range(2, ws.max_row + 1):
        user_id = ws.cell(i, column=3).value
        status = ws.cell(i, column=5).value
        if status != "Отказался":
            try:
                bot.send_message(user_id, text)
                sent += 1
            except:
                continue
    bot.send_message(message.chat.id, f"Рассылка отправлена {sent} пользователям.")

# === Запуск бота ===
print("✅ Бот запущен. Ждём сообщений...")
bot.infinity_polling()


