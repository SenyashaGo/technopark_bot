from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from imaplib import Commands
import logging
import os
from aiogram import Bot, types
from aiogram.utils.markdown import text, bold, italic, code, pre, hitalic, escape_md, _join, hbold, hcode, hpre, underline, hunderline, strikethrough, hstrikethrough, link, hlink, hide_link
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
from aiogram.types import InputFile
from aiogram.types import ReplyKeyboardRemove, \
    ReplyKeyboardMarkup, KeyboardButton, \
    InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.types import ParseMode, InputMediaPhoto, InputMediaVideo, ChatActions
from aiogram.utils import executor
from aiogram.dispatcher import Dispatcher
from aiogram.types.message import ContentType
import logging
from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.dispatcher import FSMContext
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters import Text
import aiogram.utils.markdown as fmt
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
import os.path


bot = Bot(token='5591273696:AAEWU07CR_9sc0_O9hhL6UlM-_Z5ClhE6D0')
dp = Dispatcher(bot, storage=MemoryStorage())
admin_chat_id = -804296035

logging.basicConfig(level=logging.INFO)

start_button = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопка старта
info = types.KeyboardButton("Отправить фото")
start_button.add(info)


cancel_button = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопка отмены
cancel = types.KeyboardButton('Отмена')
cancel_button.add(cancel)


ok_button2 = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопка готовности
ok2 = types.KeyboardButton('Все верно')
ok3 = types.KeyboardButton('Отмена')
ok_button2.add(ok2)
ok_button2.add(ok3)

ok_button = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопка готовности
ok = types.KeyboardButton('Готово')
ok1 = types.KeyboardButton('Отмена')
ok_button.add(ok)
ok_button.add(ok1)


date_button = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки аудиторий Пром
d1 = types.KeyboardButton("31 октября")          
d2 = types.KeyboardButton("1 ноября")
d3 = types.KeyboardButton("2 ноября")
d4 = types.KeyboardButton("3 ноября")
d5 = types.KeyboardButton("7 ноября")
d6 = types.KeyboardButton("8 ноября")
d7 = types.KeyboardButton("9 ноября")          
d8 = types.KeyboardButton("10 ноября")
d9 = types.KeyboardButton("11 ноября")
d10 = types.KeyboardButton("12 ноября")
d11 = types.KeyboardButton("14 ноября")
d12 = types.KeyboardButton("15 ноября")
d13 = types.KeyboardButton("16 ноября")          
d14 = types.KeyboardButton("17 ноября")
d15 = types.KeyboardButton("18 ноября")
d16 = types.KeyboardButton("19 ноября")
back0 = types.KeyboardButton("Отмена")






date_button.add(d16)
date_button.add(back0)


frame_button = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки корпусов
first = types.KeyboardButton("ПРОМ")          
second = types.KeyboardButton("СЕРВИС")
third = types.KeyboardButton("УРБАН")
fourth = types.KeyboardButton("ЦИФРА")
fifth = types.KeyboardButton("АРТ")
back01 = types.KeyboardButton("Отмена")
frame_button.add(first)
frame_button.add(second)
frame_button.add(third)
frame_button.add(fourth)
frame_button.add(fifth)
frame_button.add(back01)


audience_button_prom = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки аудиторий Пром
first1 = types.KeyboardButton("Аудит. 101")          
second1 = types.KeyboardButton("Аудит. 102")
third1 = types.KeyboardButton("Аудит. 108")
fourth1 = types.KeyboardButton("Аудит. 111")
fifth1 = types.KeyboardButton("Аудит. 118")
back1 = types.KeyboardButton("Отмена")
audience_button_prom.add(first1)
audience_button_prom.add(second1)
audience_button_prom.add(third1)
audience_button_prom.add(fourth1)
audience_button_prom.add(fifth1)
audience_button_prom.add(back1)


audience_button_servis = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки аудиторий Сервис
first2 = types.KeyboardButton("Аудит. 103")
back2 = types.KeyboardButton("Отмена")         
audience_button_servis.add(first2)
audience_button_servis.add(back2)


audience_button_prof = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки аудиторий Сервис
first3 = types.KeyboardButton("1 профпроба")
second3 = types.KeyboardButton("2 профпроба")
back9 = types.KeyboardButton("Отмена")       
audience_button_prof.add(first3)
audience_button_prof.add(second3)
audience_button_prof.add(back9)


audience_button_urban = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки аудиторий Урбан
first3 = types.KeyboardButton("Аудит. 105")
second3 = types.KeyboardButton("Аудит. 106")
third3 = types.KeyboardButton("Аудит. 109")
fourth3 = types.KeyboardButton("Аудит. 123")
fifth3 = types.KeyboardButton("201 ХОЛЛ")
sixth3 = types.KeyboardButton("Аудит. 201.2")
seventh12 = types.KeyboardButton("Аудит. 201.3")
eighth3 = types.KeyboardButton("Аудит. 202.1")
ninth3 = types.KeyboardButton("Аудит. 204.2")
tenth3 = types.KeyboardButton("Аудит. 209")
eleventh3 = types.KeyboardButton("Аудит. 301.1")
twelfth3 = types.KeyboardButton("Аудит. 301.3")
thirteenth3 = types.KeyboardButton("Аудит. 303.1")
fourteenth3 = types.KeyboardButton("Аудит. 303.2")
third12 = types.KeyboardButton("304 ХОЛЛ")
fifteenth3 = types.KeyboardButton("Аудит. 304.1")
sixteenth3 = types.KeyboardButton("Аудит. 304.2")
seventeenth3 = types.KeyboardButton("Аудит. 310")
back3 = types.KeyboardButton("Отмена")
audience_button_urban.add(first3)
audience_button_urban.add(second3)
audience_button_urban.add(third3)
audience_button_urban.add(fourth3)
audience_button_urban.add(fifth3)
audience_button_urban.add(sixth3)
audience_button_urban.add(seventh12)
audience_button_urban.add(eighth3)
audience_button_urban.add(ninth3)
audience_button_urban.add(tenth3)
audience_button_urban.add(eleventh3)
audience_button_urban.add(twelfth3)
audience_button_urban.add(thirteenth3)
audience_button_urban.add(fourteenth3)
audience_button_urban.add(third12)
audience_button_urban.add(fifteenth3)
audience_button_urban.add(sixteenth3)
audience_button_urban.add(seventeenth3)
audience_button_urban.add(back3)


audience_button_cifra = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки аудиторий Цифра
first4 = types.KeyboardButton("Аудит. 101")          
second4 = types.KeyboardButton("Аудит. 102")
third4 = types.KeyboardButton("Аудит. 202")
fourth4 = types.KeyboardButton("Аудит. 201.4")
fifth4 = types.KeyboardButton("Аудит. 227")
sixth4 = types.KeyboardButton("Аудит. 304")
back4 = types.KeyboardButton("Отмена")
audience_button_cifra.add(first4)
audience_button_cifra.add(second4)
audience_button_cifra.add(third4)
audience_button_cifra.add(fourth4)
audience_button_cifra.add(fifth4)
audience_button_cifra.add(sixth4)
audience_button_cifra.add(back4)


audience_button_art = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки аудиторий Арт
first5 = types.KeyboardButton("Аудит. 101")
second5 = types.KeyboardButton("Аудит. 104")
third5 = types.KeyboardButton("Аудит. 107")
fourth5 = types.KeyboardButton("108 Конф.-зал")
fifth5 = types.KeyboardButton("Аудит. 109.1")
sixth5 = types.KeyboardButton("Аудит. 204.2")
seventh5 = types.KeyboardButton("Аудит. 205.2")
eighth5 = types.KeyboardButton("Аудит. 206")
ninth5 = types.KeyboardButton("Аудит. 209.4")
tenth5 = types.KeyboardButton("Аудит. 210")
tenth9 = types.KeyboardButton("Аудит. 212")
eleventh5 = types.KeyboardButton("Аудит. 213")
twelfth5 = types.KeyboardButton("Аудит. 214")
thirteenth5 = types.KeyboardButton("ХОЛЛ BlackBox большой")
fourteenth5 = types.KeyboardButton("ХОЛЛ BlackBox маленький")
back5 = types.KeyboardButton("Отмена")
audience_button_art.add(first5)
audience_button_art.add(second5)
audience_button_art.add(third5)
audience_button_art.add(fourth5)
audience_button_art.add(fifth5)
audience_button_art.add(sixth5)
audience_button_art.add(seventh5)
audience_button_art.add(eighth5)
audience_button_art.add(ninth5)
audience_button_art.add(tenth5)
audience_button_art.add(tenth9)
audience_button_art.add(eleventh5)
audience_button_art.add(twelfth5)
audience_button_art.add(thirteenth5)
audience_button_art.add(fourteenth5)
audience_button_art.add(back5)


shift_button = types.ReplyKeyboardMarkup(resize_keyboard=True) # кнопки Смен
first6 = types.KeyboardButton("1 смена")
second6 = types.KeyboardButton("2 смена")
third6 = types.KeyboardButton("3 смена") 
back6 = types.KeyboardButton("Отмена")
shift_button.add(first6)
shift_button.add(second6)
shift_button.add(third6)
shift_button.add(back6)


class meinfo(StatesGroup):
    Q1 = State()
    Q2 = State()
    Q3 = State()
    Q4 = State()
    Q5 = State()
    Q6 = State()
    Q7 = State()


@dp.message_handler(commands='start')
async def start(message: types.Message):
    with open("keys","r") as JF:
        J = set ()
        for line in JF:
            J.add(line.strip())
        if not (str(message.chat.id)) in J:
            text = bold("ВВЕДИТЕ КОД ДОСТУПА")
            await message.answer(text, parse_mode=ParseMode.MARKDOWN)
        else:
            await message.answer(f'Привет, {message.from_user.first_name}!\nВыбери команду /photo', reply_markup=start_button)


@dp.message_handler(Text(equals="Отмена"), state="*")
async def menu_button(message: types.Message, state: FSMContext):
    await state.finish()
    await bot.send_message(message.chat.id, "Отмена произошла успешно.")
    await message.answer(f'Привет, {message.from_user.first_name}!\nВыбери команду /photo', reply_markup=start_button)


@dp.message_handler(Text(equals="Готово"), state="*")
async def menu_button(message: types.Message, state: FSMContext):  
    await state.finish()
    await bot.send_message(message.chat.id, "Фотографии сохранены.")
    await message.answer(f'Привет, {message.from_user.first_name}!\nВыбери команду /photo', reply_markup=start_button)


@dp.message_handler(commands='photo', state=None)        
async def enter_meinfo(message: types.Message):
    await message.answer("Выберите дату", reply_markup=date_button)
    await meinfo.Q1.set()                                     


@dp.message_handler(state=meinfo.Q1)                                
async def answer_q1(message: types.Message, state: FSMContext):
    answer = message.text
    await state.update_data(answer1=answer)             
    await message.answer("Выберите корпус", reply_markup=frame_button)
    await meinfo.Q2.set()  
 

@dp.message_handler(state=meinfo.Q2)                                
async def answer_q1(message: types.Message, state: FSMContext):
    answer = message.text
    await state.update_data(answer2=answer)
    if answer == 'ПРОМ':
        await message.answer("Выберите аудиторию", reply_markup=audience_button_prom)
    elif answer == 'СЕРВИС':
        await message.answer("Выберите аудиторию", reply_markup=audience_button_servis)
    elif answer == 'УРБАН':
        await message.answer("Выберите аудиторию", reply_markup=audience_button_urban)
    elif answer == 'ЦИФРА':
        await message.answer("Выберите аудиторию", reply_markup=audience_button_cifra)
    elif answer == 'АРТ':
        await message.answer("Выберите аудиторию", reply_markup=audience_button_art)
    await meinfo.Q3.set()                                   

@dp.message_handler(state=meinfo.Q3)                                
async def answer_q1(message: types.Message, state: FSMContext):
    answer = message.text
    await state.update_data(answer3=answer)                            
    await message.answer("Выберите смену", reply_markup=shift_button)
    await meinfo.Q4.set()


@dp.message_handler(state=meinfo.Q4)                                
async def answer_q1(message: types.Message, state: FSMContext):
    answer = message.text
    await state.update_data(answer4=answer)                            
    await message.answer("Выберите профпробу", reply_markup=audience_button_prof)
    await meinfo.Q5.set()


@dp.message_handler(state=meinfo.Q5)                                
async def answer_q1(message: types.Message, state: FSMContext):
    answer = message.text
    await state.update_data(answer5=answer)   
    data = await state.get_data()               
    answer1 = data.get("answer1")               
    answer2 = data.get("answer2")
    answer3 = data.get("answer3")
    answer4 = data.get("answer4") 
    answer5 = data.get("answer5")                  
    await message.answer(f"Вы выбрали: {answer1}, {answer2}, {answer3}, {answer4}, {answer5}", reply_markup=ok_button2)
    await meinfo.Q6.set()

@dp.message_handler(state=meinfo.Q6)                                
async def answer_q1(message: types.Message, state: FSMContext):
    answer = message.text
    data = await state.get_data()               
    answer1 = data.get("answer1")               
    answer2 = data.get("answer2")
    answer3 = data.get("answer3")
    answer4 = data.get("answer4")   
    answer5 = data.get("answer5")
    if answer == "Все верно":
        await message.answer("Прикрепите фото", reply_markup=ok_button)
        await bot.send_message(admin_chat_id, f"{answer1}/{answer2}/{answer3}/{answer4}/{answer5}")
        await meinfo.Q7.set()
    else:
        await state.finish()
        await message.answer(f'Привет, {message.from_user.first_name}!\nВыбери команду /photo', reply_markup=start_button)



@dp.message_handler(content_types=["photo"], state=meinfo.Q7)
async def photo_handler(message: types.Message, state: FSMContext):
    data = await state.get_data()               
    answer1 = data.get("answer1")               
    answer2 = data.get("answer2")
    answer3 = data.get("answer3")
    answer4 = data.get("answer4")
    answer5 = data.get("answer5")
    if os.path.exists(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}") == False:
        os.makedirs(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}")
    if os.path.exists(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}") == False:
        os.makedirs(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}")
    if os.path.exists(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}//{answer3}") == False:
        os.makedirs(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}//{answer3}")
    if os.path.exists(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}//{answer3}//{answer4}") == False:
        os.makedirs(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}//{answer3}//{answer4}")
    if os.path.exists(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}//{answer3}//{answer4}//{answer5}") == False:
        os.makedirs(f"C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}//{answer3}//{answer4}//{answer5}")

    photo = message.photo.pop()
    await photo.download(f'C://Users//Administrator//technopark_bot//ФОТО//{answer1}//{answer2}//{answer3}//{answer4}//{answer5}//')
    # await bot.send_photo(admin_chat_id, photo=message.photo[-1].file_id)
    



@dp.message_handler(content_types=['text'], state=None)
async def enter_meinfo(message: types.Message):
    if message.text == '666':
        with open("keys","r") as JF:
            J = set ()
            for line in JF:
                J.add(line.strip())
            JF = open("keys","a")
            JF.write(str(message.chat.id) + '\n')
            J.add(message.chat.id)
            await message.answer(f'Привет, {message.from_user.first_name}!\nВыбери команду /photo', reply_markup=start_button)
    
    elif message.text == 'Отправить фото' or message.text == 'отправить фото':
        await message.answer("Выберите дату", reply_markup=date_button)
        await meinfo.Q1.set()
    # elif int(message.chat.id) == int(admin_chat_id):
    #     chat_text = message.text.split(': ')[0]
    #     if chat_text == 'выгрузка' or chat_text == 'Выгрузка':
    #         mesto = "/Users/senyashago/Desktop/PythonWork/Бот мамина работа/" + message.text.split(': ')[1]
    #         for filename in os.listdir('/Users/senyashago/Desktop/PythonWork/Бот мамина работа/'):
    #             await bot.send_photo(admin_chat_id, photo=filename)
    # for filename in os.listdir('/Users/senyashago/Desktop/PythonWork/Бот мамина работа/2 ноября/ЦИФРА/Аудит. 201.4/3 смена/photos/'):
    #     await bot.send_photo(admin_chat_id, photo=open(f'/Users/senyashago/Desktop/PythonWork/Бот мамина работа/2 ноября/ЦИФРА/Аудит. 201.4/3 смена/photos/{filename}', 'rb'))

    elif int(message.chat.id) == int(admin_chat_id):
        chat_text = message.text.split(': ')[0]
        if chat_text == 'выгрузка' or chat_text == 'Выгрузка':
            mesto = "C://Users//Administrator//technopark_bot//ФОТО//" + message.text.split(': ')[1] + '/photos/'
            for filename in os.listdir(mesto):
                await bot.send_photo(admin_chat_id, photo=open(f'{mesto}{filename}', 'rb'))




if __name__ == '__main__':
    executor.start_polling(dp)