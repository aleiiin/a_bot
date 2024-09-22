import configparser
import contextlib

from xlsxwriter.workbook import Workbook
from db import *
import logging
from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import (Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, ReplyKeyboardMarkup,
                           ReplyKeyboardRemove)
from datetime import datetime
from apscheduler.schedulers.asyncio import AsyncIOScheduler

menu_kb = ReplyKeyboardMarkup(resize_keyboard=True)
menu_kb.add(InlineKeyboardButton(text="🥰Тарифы🥰", callback_data='tariffs'))
menu_kb.add(InlineKeyboardButton(text="🌟Мой профиль🌟", callback_data='my_profile'))
menu_kb.add(InlineKeyboardButton(text="📞Обратная связь📞", callback_data='callback'))
menu_kb.add(InlineKeyboardButton(text="🔥Предложить себя🔥", callback_data='offer_me'))
menu_kb.add(InlineKeyboardButton(text="Админ меню", callback_data='show_menu_adm'))

menu_user = ReplyKeyboardMarkup(resize_keyboard=True)
menu_user.add(InlineKeyboardButton(text="🥰Тарифы🥰", callback_data='tariffs'))
menu_user.add(InlineKeyboardButton(text="🌟Мой профиль🌟", callback_data='my_profile'))
menu_user.add(InlineKeyboardButton(text="📞Обратная связь📞", callback_data='callback'))
menu_user.add(InlineKeyboardButton(text="🔥Предложить себя🔥", callback_data='offer_me'))


menu_adm = InlineKeyboardMarkup(resize_keyboard=True)
menu_adm.add(InlineKeyboardButton(text='Редактировать тарифы', callback_data='edit_tariffs'))
menu_adm.add(InlineKeyboardButton(text='Рассылка', callback_data='mailing'))
menu_adm.add(InlineKeyboardButton(text='Изменить готовую фразу', callback_data='new_phrase'))
menu_adm.add(InlineKeyboardButton(text='Оплата и пополнение', callback_data='edit_deposit'))
menu_adm.add(InlineKeyboardButton(text="Статистика", callback_data='stat'))
menu_adm.add(InlineKeyboardButton(text="Найти пользователя по ID", callback_data='search'))
menu_adm.add(InlineKeyboardButton(text="Контакт тех. поддержки", callback_data='edit_helper'))
menu_adm.add(InlineKeyboardButton(text="Администраторы", callback_data='admin_list'))
menu_adm.add(InlineKeyboardButton(text="Предложения", callback_data='offers'))

logging.basicConfig(level=logging.INFO)

storage = MemoryStorage()
# bot = Bot(token=API_TOKEN, proxy='http://proxy.server:3128')
# "http://QDP7WVW8O1:XMFDDihuPd@194.31.73.156:31739"

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot, storage=storage)

cur = con.cursor()

cur.execute('SELECT * FROM Settings')
if not cur.fetchone():
    cur.execute('INSERT INTO Settings VALUES (?, ?, ?, ?, ?)', (1, 1, '', '', 0))
for admin in ADMINS:
    cur.execute(f'SELECT * FROM Admins WHERE user_id={admin}')
    if not cur.fetchone():
        cur.execute('INSERT INTO Admins VALUES (?, ?, ?, ?, ?)', (admin, get_date(), 1, 1, 1))
    cur.execute(f'SELECT * FROM Users WHERE user_id={admin}')
    if not cur.fetchone():
        cur.execute('INSERT INTO Users VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)', (admin, '', 0, get_date(), 0, 0, 1, 0, 0))

try:
    file = open('tariffs.txt', 'r', encoding='UTF-8')
except:
    print('Проблема с открытием файла с тарифами')
else:
    for line in file.readlines():
        res = list(map(lambda x: x.strip(), line.split('/')))
        if len(res) in (4, 5):
            if res[1] == '0':
                res[1] = 'Навсегда'
            if len(res) == 4:
                res.insert(2, '')
            name = res[0]
            cur.execute(f'SELECT * FROM Tariffs WHERE work=1 AND name="{name}" AND days="{res[1]}"')
            if not cur.fetchone():
                try:
                    cur.execute('INSERT INTO Tariffs VALUES (?, ?, ?, ?, ?)', (name, res[1],
                                                                       res[2], res[3], 1))
                except:
                    print(f'Ошибка с добавлением канала {res}')
                else:
                    con.commit()
            link = res[4]
            if link[1:].isdigit():
                cur.execute(f'SELECT * FROM Tariffs_links WHERE name="{name}"')
                if not cur.fetchone():
                    try:
                        cur.execute('INSERT INTO Tariffs_links VALUES (?, ?)', (name, link))
                    except:
                        print(f'Ошибка с добавлением айди канала {res}')
                    else:
                        con.commit()
                else:
                    try:
                        cur.execute(f'UPDATE Tariffs_links SET id={link} WHERE name="{name}"')
                    except:
                        print(f'Ошибка с обновлением айди канала {res}')
                    else:
                        con.commit()
        elif len(res) == 2:
            name, number = res
            try:
                cur.execute(f'SELECT * FROM Payment_methods WHERE name="{name}"')
            except:
                print(f'Проблема с добавлением способа оплаты {res}')
            else:
                if not cur.fetchone():
                    try:
                        cur.execute('INSERT INTO Payment_methods VALUES (?, ?, ?, ?, ?, ?)', (name,
                                                                                               '', 0.0, number, 0,
                                                                                               1))
                    except:
                        print(f'Проблема с добавлением способа оплаты {res}')
                    else:
                        con.commit()
                else:
                    try:
                        cur.execute(f'UPDATE Payment_methods SET number={number} WHERE name="{name}"')
                    except:
                        print(f'Проблема с обновлением способа оплаты {res}')
                    else:
                        con.commit()
        # elif len(res) == 1:
        #     try:
        #         cur.execute('SELECT * FROM Tariffs WHERE name="Всё вместе навсегда"')
        #     except:
        #         print('Проблема с тарифом "Всё вместе навсегда"')
        #     else:
        #         if not cur.fetchone():
        #             try:
        #                 cur.execute('INSERT INTO Tariffs VALUES (?, ?, ?, ?, ?)', ('Всё вместе навсегда', 'Навсегда', '',
        #                                                                            res[0], 1))
        #             except:
        #                 print('Проблема с добавлением тарифа "Всё вместе навсегда"')
        #             else:
        #                 con.commit()

try:
    file = open('ban.txt', 'r', encoding='UTF-8')
except:
    print('Проблема с открытием файла с забаненными')
else:
    for line in file.readlines():
        try:
            user = int(line.strip())
        except:
            print(f'Ошибка с баном {line}')
        else:
            cur.execute(f'SELECT * FROM Users WHERE user_id={user}')
            if cur.fetchone():
                cur.execute(f'UPDATE Users SET ban=1 WHERE user_id={user}')
                con.commit()

scheduler = AsyncIOScheduler()


async def check():
    cur.execute('SELECT * FROM Admins')
    admins = tuple(map(lambda x: x[0], cur.fetchall()))
    cur.execute('SELECT * FROM Subs WHERE days != "Навсегда"')
    subs = cur.fetchall()
    for i in range(len(subs)):
        user_id, days, tariff, date = subs[i]
        tariff = tariff.split('/')[0]
        buy_date = date
        date = str(date).split('/')
        date_now = str(get_date()).split('/')
        day1, month1, year1 = tuple(map(int, list(date_now)))
        day2, month2, year2 = tuple(map(int, list(date)))
        date_now = datetime(year1, month1, day1)
        date = datetime(year2, month2, day2)
        delta = str(date_now - date)
        try:
            delta = int(delta.split(',')[0].split()[0])
        except:
            for admin in admins:
                await bot.send_message(chat_id=admin, text=f'Произошла ошибка с удалением из канала: {tariff}\n'
                                                           f'User id: {user_id}\nТариф: {tariff}/{days}\nДата покупки'
                                                           f': {buy_date}')
        else:
            cur.execute(f'SELECT * FROM Tariffs_links WHERE name="{tariff}"')
            group_id = cur.fetchone()
            if group_id:
                group_id = group_id[1]
                try:
                    status = await bot.get_chat_member(chat_id=group_id, user_id=user_id)
                    status = status['status']
                except:
                    for admin in admins:
                        await bot.send_message(chat_id=admin, text=f'❗Внимание❗\nУ бота проблемы с каналом {tariff}\n'
                                                                   f'Возможно id канала указан неверно или у бота недостаточно '
                                                                   f'прав администратора в канале')
                else:
                    if status == 'member':
                        if delta > days:
                            await bot.send_message(chat_id=user_id, text=f'❗Внимание❗\n'
                                                                         f'У вас закончился срок подписки на канал {tariff.split("/")}\n'
                                                                         f'Вы исключены из данного канала')
                            try:
                                await bot.kick_chat_member(chat_id=int(group_id), user_id=user_id)
                            except:
                                for admin in admins:
                                    await bot.send_message(chat_id=admin, text=f'❗Внимание❗\nНе удалось исключить  пользователя '
                                                                               f'{user_id} из канала {tariff}')
                            else:
                                print(f'{user_id} исключён из канала {tariff}')
                                cur.execute(f'DELETE FROM Subs WHERE user_id={user_id} AND tariff="{tariff}"')
                                con.commit()
                            cur.execute(f'SELECT * FROM Subs WHERE user_id={user_id}')
                            if tuple(cur.fetchall()):
                                pass
                            else:
                                cur.execute(f'UPDATE Users SET buy=0 WHERE user_id={user_id}')
                                con.commit()
                        elif delta == days:
                            await bot.send_message(chat_id=user_id, text=f'❗Внимание❗\n'
                                                                   f'Завтра заканчивается срок подписки на канал {tariff}')
            else:
                try:
                    for admin in admins:
                        await bot.send_message(chat_id=admin, text=f'Произошла ошибка с id канала {tariff}')
                except:
                    print(f'Произошла ошибка с id канала {tariff}')
                else:
                    pass


async def schedule_jobs(_):
    scheduler.add_job(check, trigger='interval', hours=24)
    scheduler.start()


def db_find_user(id):
    id = int(id)
    user = [None, None]
    cur.execute(f'SELECT * FROM Users WHERE user_id={id}')
    user[0] = tuple(cur.fetchone())
    if user[0][1] == 1:
        cur.execute(f'SELECT * FROM Subs WHERE user_id = {id}')
        user[1] = tuple(cur.fetchone())

    return user


class Dialog(StatesGroup):

    user_id = State()

    check_new_tariff_state = State()
    add_new_tariff_state = State()
    choice_time_state = State()
    choice_pay_method_state = State()
    accept_pay_state = State()
    pay_fiat_state = State()
    check_pay_fiat_state = State()

    choice_to_delete_state = State()
    ans_to_delete_state = State()
    delete_tariffs_state = State()
    check_tariffs_links_state = State()

    mailing_state = State()
    mailing_new_tariff_state = State()
    mailing_start_state = State()

    mailing_text = State()
    mailing_img = State()
    mailing_caption = State()

    edit_tariff_state = State()
    edit_tariff_choice_state = State()
    edit_tariff_start_state = State()
    edit_tariff_ans_state = State()
    add_tariff_always_state = State()
    check_tariff_always_state = State()

    buyer_photo = State()
    accept_purchase_state = State()
    cancel_purchase_state = State()

    spam_state = State()
    choice_spam_state = State()
    send_spam_state = State()

    edit_phrase_state = State()

    edit_helper_state = State()

    choice_tariff = State()
    choice_days = State()
    choice_description = State()
    choice_price = State()
    choice_method = State()
    choice_promo = State()

    deposit_state = State()

    edit_min_deposit_state = State()

    edit_payment_method_state = State()
    choice_edit_payment_method_state = State()
    edit_name_state = State()
    edit_name_ans_state = State()
    edit_requisites_state = State()
    choice_edit_requisites_state = State()
    edit_requisites_ans_state = State()

    new_method = State()
    old_requisites = State()
    new_requisites = State()

    add_payment_method_state = State()
    add_payment_method_ans_state = State()

    delete_payment_method_state = State()

    edit_forever_state = State()

    balance_pay_check_state = State()

    balance = State()

    deposit_buy = State()
    choice_pay_method = State()
    columns = State()

    helper = State()

    offer_me_state = State()
    media = State()
    caption = State()
    type = State()
    ans_offer_me_state = State()


@dp.message_handler(commands=['start'])
async def start(msg: Message):
    user = msg.from_user.id
    cur.execute(f"SELECT user_id FROM Users WHERE user_id = {user}")
    result = cur.fetchone()
    if result is None:
        date = get_date()
        if user in ADMINS:
            cur.execute("INSERT INTO Users VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?);", (user, msg.from_user.username,
                                                                           0, date, 0, 0, 1, 0, 0))
            cur.execute("INSERT INTO Admins VALUES(?, ?, ?, ?, ?);", (user, date, 1, 1, 0))
        else:
            cur.execute("INSERT INTO Users VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?);", (user, msg.from_user.username,
                                                                           0, date, 0, 0, 0, 0, 0))
    if user in ADMINS:
        await msg.answer(f'{msg.from_user.full_name} добро пожаловать в бота!', reply_markup=menu_kb)
    else:
        await msg.answer(f'{msg.from_user.full_name} добро пожаловать в бота!', reply_markup=menu_user)

    con.commit()


@dp.message_handler(text='Админ меню')
async def admin_btn(msg: Message):
    user_id = msg.from_user.id
    cur.execute(f'SELECT * FROM Admins WHERE user_id={user_id}')
    admin = cur.fetchone()
    if admin:
        await msg.answer('Добро пожаловать в админ меню', reply_markup=menu_kb)
        await msg.answer(f'Ваш id:{admin[0]}\nДата назначения админом: {admin[1]}\n\nВыберите действие ',
                         reply_markup=menu_adm)
    else:
        await msg.answer('Вас нет в списке админов', reply_markup=menu_user)


@dp.callback_query_handler(text='edit_tariffs')
async def show_stat(call: CallbackQuery):
    await call.message.delete()

    kb = InlineKeyboardMarkup()
    kb.add(InlineKeyboardButton('Добавить тарифы', callback_data='add_tariffs'))
    kb.add(InlineKeyboardButton('Удалить тариф', callback_data='delete_tariffs'))
    kb.add(InlineKeyboardButton('Редактировать существующие тарифы', callback_data='edit_tariff'))
    kb.add(InlineKeyboardButton('Редактировать "Всё вместе навсегда"', callback_data='edit_forever'))
    await call.message.answer('Выберите действие', reply_markup=kb)


@dp.callback_query_handler(text='add_tariffs')
async def add_tariffs_func(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    if call.data == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        kb = ReplyKeyboardMarkup()
        kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
        await call.message.answer('Введите данные о новом тарифе в виде:\n\nназвание/кол-во дней'
                                  '/описание (не обязательно)/цена/id группы (если нет в базе данных)\n\n'
                                  '"Отмена" для отмены', reply_markup=kb)
        await Dialog.check_new_tariff_state.set()


@dp.message_handler(state=Dialog.check_new_tariff_state)
async def check_new_tariff_set(msg: Message, state=FSMContext):
    tariff = msg.text.split('/')
    if tariff[0].lower() == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    elif len(tariff) != 5:
        await msg.answer('В введённых вами данных какая-то ошибка. Попробуйте заново', reply_markup=menu_adm)
        await state.finish()
    else:
        try:
            tariff[0] = tariff[0].capitalize()
        except:
            await msg.answer('В введённых вами данных какая-то ошибка. Попробуйте заново', reply_markup=menu_adm)
            await state.finish()
        else:
            name = tariff[0]
            cur.execute(f'SELECT * FROM Tariffs_links WHERE name="{name}"')
            if tariff[4] and not str(tariff[4][1:]).isdigit():
                await msg.answer('ID группы введён неверно', reply_markup=menu_kb)
                await state.finish()
            else:
                await state.update_data(choice_tariff=tariff)
                kb = InlineKeyboardMarkup()
                kb.add(InlineKeyboardButton('Да', callback_data='yes'))
                kb.add(InlineKeyboardButton('Нет', callback_data='to_menu'))
                await msg.delete()
                await msg.answer(f'Название: {name}\nКол-во дней: {tariff[1]}\nОписание: {tariff[2]}\n'
                                 f'Цена: {tariff[3]} ₽\n\nДобавить тариф?', reply_markup=kb)
                await Dialog.add_new_tariff_state.set()


@dp.callback_query_handler(state=Dialog.add_new_tariff_state)
async def add_new_tariff_set(call: CallbackQuery, state=FSMContext):
    if call.data == 'yes':
        data = await state.get_data()
        tariff = data.get('choice_tariff')
        name = tariff[0]
        cur.execute(f'SELECT * FROM Tariffs WHERE name="{name}" AND days="{tariff[1]}"')
        if cur.fetchall():
            await call.message.answer(f'Данный тариф ({name}/{tariff[1]} дней) уже существует', reply_markup=menu_kb)
            await state.finish()
        else:
            cur.execute(f'INSERT INTO Tariffs VALUES (?, ?, ?, ?, ?);', (name, tariff[1], tariff[2], tariff[3], 1))
            cur.execute('SELECT * FROM Tariffs_links')
            if name not in tuple(map(lambda x: x[0], cur.fetchall())):
                cur.execute(f'INSERT INTO Tariffs_links VALUES (?, ?)', (name, tariff[4]))
            else:
                cur.execute(f'SELECT * FROM Tariffs_links WHERE name={name}')
                await call.message.answer(f'ID тарифа {name} изменён с {cur.fetchone()[1]} на {tariff[4]}')
                cur.execute(f'UPDATE Tariffs_links SET id={tariff[4]}')
            con.commit()
            cur.execute(f'SELECT * FROM Tariffs WHERE name = "{name}" AND work = 1 AND days="Навсегда"')
            if cur.fetchone():
                kb = InlineKeyboardMarkup()
                kb.add(InlineKeyboardButton('Да', callback_data='mailing'))
                kb.add(InlineKeyboardButton('Нет', callback_data='to_menu'))
                await call.message.answer(f'Тариф {name}/{tariff[1]} успешно добавлен\nСделать рассылку?',
                                          reply_markup=menu_kb)
                await state.finish()
            else:
                await call.message.answer('Введите цену для тарифа "Навсегда"')
                await Dialog.add_tariff_always_state.set()
    else:
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()


@dp.message_handler(state=Dialog.add_tariff_always_state)
async def add_tariff_always_set(msg: Message, state=FSMContext):
    price = msg.text
    try:
        price = int(price)
    except:
        new_price = ''
        for x in price:
            if x.isdigit():
                new_price += x
        await state.update_data(choice_price=new_price)
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton('Да', callback_data='yes'))
        kb.add(InlineKeyboardButton('Нет', callback_data='to_menu'))
        await msg.answer(f'Возможно вы имелии ввиду "{new_price}"?', reply_markup=kb)
        await Dialog.check_tariff_always_state.set()
    else:
        data = await state.get_data()
        tariff = data.get('choice_tariff')
        name = tariff[0]
        try:
            cur.execute(f'INSERT INTO Tariffs VALUES (?, ?, ?, ?, ?);', (name, 'Навсегда', tariff[2], price, 1))
        except:
            await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()
        else:
            con.commit()
            kb = InlineKeyboardMarkup()
            kb.add(InlineKeyboardButton('Да', callback_data='mailing'))
            kb.add(InlineKeyboardButton('Нет', callback_data='to_menu'))
            await msg.answer(f'Тариф "{name}" успешно добавлен\nСделать рассылку?', reply_markup=kb)
            await Dialog.mailing_new_tariff_state.set()


@dp.callback_query_handler(state=Dialog.check_tariff_always_state)
async def check_tariff_always_set(call: CallbackQuery, state=FSMContext):
    pick = call.data
    data = await state.get_data()
    tariff = data.get("choice_tariff")
    if pick == 'to_menu':
        await call.message.answer(f'Вы вернулись в меню\nДля добавления тарифа "Навсегда" в группу '
                                  f'{tariff[0]} перейдите в:\nАдмин меню -> Редактировать тарифы ->'
                                  f'Редактировать существующие тарифы', reply_markup=menu_kb)
        await state.finish()
    elif pick == 'yes':

        cur.execute(f'INSERT INTO Tariffs VALUES (?, ?, ?, ?, ?);', (tariff[0], 'Навсегда', tariff[2],
                                                                     data.get('choice_price'), 1))
        con.commit()
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton('Да', callback_data='mailing'))
        kb.add(InlineKeyboardButton('Нет', callback_data='to_menu'))
        await call.message.answer('Тариф успешно добавлен\nСделать рассылку?', reply_markup=kb)
        await Dialog.mailing_new_tariff_state.set()
    else:
        await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)


@dp.callback_query_handler(text='mailing')
async def mailing_func(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    if call.data == 'to_menu':
        await call.message.answer('Тариф успешно добавлен. Рассылка отменена', reply_markup=menu_kb)
    else:
        kb = ReplyKeyboardMarkup()
        kb.add(InlineKeyboardButton('Готовая фраза', callback_data='phrase'))
        kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
        cur.execute('SELECT * FROM Settings')
        phrase = cur.fetchone()[2]
        if not phrase:
            phrase = 'None'
        await state.update_data(mailing_text=phrase)
        await call.message.answer(f'Отправьте изображение / текст для рассылки\nГотовая фраза: {phrase}', reply_markup=kb)
        await Dialog.mailing_state.set()


@dp.message_handler(content_types=['photo', 'text'], state=Dialog.mailing_state)
async def mailing_set(msg: Message, state=FSMContext):
    pick = msg.photo
    kb = InlineKeyboardMarkup()
    kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    if pick:
        caption = ''
        pick = pick[0].file_id
        if msg['caption']:
            caption = msg['caption']
        await state.update_data(mailing_text=caption)
        await state.update_data(mailing_img=pick)
        await msg.answer('Рассылка:')
        await bot.send_photo(chat_id=msg.from_user.id, photo=pick, caption=caption, reply_markup=kb)
        await Dialog.mailing_start_state.set()
    elif msg.text.lower() == 'готовая фраза':
        data = await state.get_data()
        if not data.get('choice_tariff'):
            await state.update_data(choice_tariff='')
        data = await state.get_data()
        text = f'{data.get("phrase")} {data.get("choice_tariff")}'
        await state.update_data(mailing_text=text)
        await msg.answer('Рассылка:')
        await bot.send_message(chat_id=msg.from_user.id, text=text, reply_markup=kb)
        await Dialog.mailing_start_state.set()
    else:
        pick = msg.text
        if pick.lower() == 'отмена':
            await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
            await state.finish()
        else:
            await state.update_data(mailing_text=pick)
            await msg.answer('Рассылка:')
            await bot.send_message(chat_id=msg.from_user.id, text=pick, reply_markup=kb)
            await Dialog.mailing_start_state.set()


@dp.callback_query_handler(state=Dialog.mailing_start_state)
async def mailing_start_set(call: CallbackQuery, state=FSMContext):
    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Рассылка отменена\nВы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    elif pick == 'yes':
        cur.execute('SELECT * FROM Users')
        users = tuple(map(lambda x: x[0], cur.fetchall()))
        data = await state.get_data()
        photo = data.get('mailing_img')
        text = data.get('mailing_text')
        length1 = len(users)
        length = len(users)
        if photo:
            for user in users:
                try:
                    await bot.send_photo(chat_id=user, photo=photo, caption=text)
                except:
                    length -= 1
        else:
            for user in users:
                try:
                    await bot.send_message(chat_id=str(user).strip(), text=text)
                except:
                    length -= 1
        await call.message.answer(f'{length}/{length1} пользователей получили рассылку', reply_markup=menu_kb)
        await state.finish()
    else:
        await call.message.answer('Рассылка отменена\nПроизошла неизвестная ошибка', reply_markup=menu_kb)
        await state.finish()


@dp.callback_query_handler(text='delete_tariffs')
async def delete_tariffs_func(call: CallbackQuery):
    await call.message.delete()

    cur.execute('SELECT * FROM Tariffs WHERE work=1 AND name != "Всё вместе навсегда"')
    res = sorted(tuple(set(map(lambda x: x[0], cur.fetchall()))))
    tariffs = ''
    for i in range(len(res)):
        tariffs += f'{i + 1}. {res[i]}\n'
    kb = ReplyKeyboardMarkup()
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    kb.add(InlineKeyboardButton('Удалить все тарифы', callback_data='delete_all_tariffs'))
    await call.message.answer('Удаление тарифов', reply_markup=ReplyKeyboardRemove())
    await call.message.answer(f'Выберите тариф или несколько тарифов (Пример: 1,2,3):\n\n{tariffs}',
                              reply_markup=kb)
    await Dialog.choice_to_delete_state.set()


@dp.message_handler(state=Dialog.choice_to_delete_state)
async def choice_to_delete_set(msg: Message, state=FSMContext):
    pick = msg.text.lower()
    if pick == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    elif pick == 'удалить все тарифы':
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
        kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
        await msg.answer('Функция временно недоступна', reply_markup=menu_kb)
    else:
        try:
            pick = tuple(map(int, pick.split(',')))
        except ValueError:
            await msg.answer('Данные введены неверно', reply_markup=menu_kb)
            await state.finish()
        except:
            await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()
        else:
            cur.execute('SELECT * FROM Tariffs WHERE work=1 AND name != "Всё вместе навсегда"')
            res = sorted(tuple(set(map(lambda x: x[0], cur.fetchall()))))
            tariffs = list()
            days = list()
            for i in range(len(pick)):
                tariffs.append(res[pick[i] - 1])
            days_to_show = ''
            count = 1
            for tariff in tariffs:
                cur.execute(f'SELECT * FROM Tariffs WHERE work=1 AND name="{tariff}"')
                res = cur.fetchall()
                for day in res:
                    days.append(day)
                    days_to_show += f'{count}. {day}\n'
                    count += 1
            await state.update_data(choice_tariff=days)
            kb = ReplyKeyboardMarkup()
            kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
            kb.add(InlineKeyboardButton('Удалить всё', callback_data='delete_all'))
            cur.execute('SELECT * FROM Tariffs')
            await state.update_data(columns=", ".join(tuple(map(lambda x: x[0], cur.description))))
            data = await state.get_data()
            await msg.answer(f'Тарифы в формате: [{data.get("columns")}]')
            await msg.answer(f'Выберите тариф или несколько тарифов (Пример: 1,2,3):\n\n{days_to_show}',
                             reply_markup=kb)
            await Dialog.ans_to_delete_state.set()


@dp.message_handler(state=Dialog.ans_to_delete_state)
async def ans_to_delete_set(msg: Message, state=FSMContext):
    pick = msg.text
    if pick.lower() == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    elif pick.lower() == 'удалить всё':
        await msg.answer('Данная функция временно недоступна', reply_markup=menu_kb)
        await state.finish()
        pass
        # kb = InlineKeyboardMarkup()
        # kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
        # kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
        # await msg.answer('Вы уверены?', reply_markup=kb)
    else:
        try:
            pick = tuple(map(int, pick.split(',')))
        except ValueError:
            await msg.answer('Данные введены неверно', reply_markup=menu_kb)
            await state.finish()
        except:
            await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()
        else:
            data = await state.get_data()
            tariffs = data.get('choice_tariff')
            tariffs_to_delete = list()
            tariffs_to_delete_show = ''
            for i in range(len(pick)):
                tariffs_to_delete.append(tariffs[pick[i] - 1])
                tariffs_to_delete_show += f'{i + 1}. {tariffs[pick[i] - 1]}\n'
            await state.update_data(choice_tariff=tariffs_to_delete)
            kb = InlineKeyboardMarkup()
            kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
            kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
            await msg.answer(f'Тарифы в формате: ({data.get("columns")})', reply_markup=ReplyKeyboardRemove())
            await msg.answer(f'Удалить тарифы:\n{tariffs_to_delete_show}', reply_markup=kb)
            await Dialog.delete_tariffs_state.set()


@dp.callback_query_handler(state=Dialog.delete_tariffs_state)
async def delete_tariffs_set(call: CallbackQuery, state=FSMContext):
    if call.data == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    elif call.data == 'yes':
        data = await state.get_data()
        tariffs_to_del = data.get('choice_tariff')
        try:
            length = len(tariffs_to_del)
        except:
            length = 0
        else:
            count = 0
            for i in range(len(tariffs_to_del)):
                try:
                    tariff = tariffs_to_del[i]
                    cur.execute(f'DELETE FROM Tariffs WHERE work = 1 AND name = "{tariff[0]}" AND '
                                f'days = "{tariff[1]}" AND price = "{tariff[3]}"')
                    con.commit()
                except:
                    await call.message.answer(f'Произошла неизвестная ошибка')
                else:
                    count += 1
            tariffs_to_del = tuple(set(map(lambda x: x[0], tariffs_to_del)))
            await state.update_data(choice_tariff=tariffs_to_del)
            kb = InlineKeyboardMarkup()
            kb.add(InlineKeyboardButton('Да', callback_data='yes'))
            kb.add(InlineKeyboardButton('Нет', callback_data='to_menu'))
            await call.message.answer(f'Удалено {count}/{length} тарифов', reply_markup=menu_kb)
            # tariffs_to_del = set(list(map(lambda x: x.split('/'), tariffs_to_del)))
            cur.execute('SELECT * FROM Tariffs WHERE work=1')
            tariffs = cur.fetchall()
            tariffs = set(list(map(lambda x: x[0], tariffs)))
            for tariff in tariffs_to_del:
                if not tariff in tariffs:
                    cur.execute(f'DELETE FROM Tariffs_links WHERE name="{tariff}"')
            con.commit()
            await state.finish()
    else:
        await call.message.answer('Произошла неизвестная ошибка.', reply_markup=menu_kb)
        await state.finish()


@dp.callback_query_handler(text='edit_tariff')
async def edit_tariff_func(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    await call.message.answer('Изменение существующих тарифов', reply_markup=ReplyKeyboardRemove())

    kb = InlineKeyboardMarkup()
    cur.execute('SELECT * FROM Tariffs WHERE work=1 AND name != "all"')
    tariffs = sorted(tuple(set(map(lambda x: x[0], cur.fetchall()))))
    for tariff in tariffs:
        kb.add(InlineKeyboardButton(tariff, callback_data=tariff))
    kb.add(InlineKeyboardButton('Группы без id / с неверным id', callback_data='not_correct'))
    kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
    await call.message.answer('Выберите групппу', reply_markup=kb)
    await Dialog.edit_tariff_state.set()


@dp.callback_query_handler(state=Dialog.edit_tariff_state)
async def edit_tariff_set(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        cur.execute('SELECT * FROM Tariffs WHERE work=1 AND name != "all"')
        tariffs = list(set(map(lambda x: x[0], cur.fetchall())))
        if pick == 'not_correct':
            cur.execute('SELECT * FROM Tariffs_links')
            links = cur.fetchall()
            for i in range(len(links)):
                tariff = links[i][0]
                if tariff in tariffs:
                    link = links[i][1]
                    if link:
                        if str(link).isdigit():
                            tariffs.remove(tariff)
            if tariffs:
                await call.message.answer(f'Группы без id / с неверным id: {tariffs}', reply_markup=menu_kb)
            else:
                await call.message.answer('У каждой группы есть корректный id')
            await state.finish()
        else:
            kb = InlineKeyboardMarkup()
            for tariff in tariffs:
                kb.add(InlineKeyboardButton(tariff, callback_data=tariff))
            kb.add(InlineKeyboardButton('Всё вместе навсегда', callback_data='Всё вместе навсегда'))
            kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
            await call.message.answer('Группы:', reply_markup=kb)
            await Dialog.edit_tariff_choice_state.set()


@dp.callback_query_handler(state=Dialog.edit_tariff_choice_state)
async def edit_tariff_choice_set(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        await state.update_data(choice_tariff=pick)
        cur.execute(f'SELECT * FROM Tariffs WHERE work=1 AND name="{pick}"')
        tariffs = cur.fetchall()
        description = tariffs[0][2]
        await state.update_data(choice_description=description)
        if pick == 'Всё вместе навсегда':
            cur.execute(f'SELECT * FROM Tariffs WHERE work=1 AND name="{pick}"')
            price = cur.fetchone()[3]
            kb = InlineKeyboardMarkup()
            kb.add(InlineKeyboardButton('Изменить цену', callback_data='change_price'))
            kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
            await call.message.answer(f'Тариф: {pick}\nЦена: {price}\nОписание: {description}', reply_markup=kb)
            await Dialog.edit_tariff_start_state.set()
        else:
            kb = InlineKeyboardMarkup()
            for tariff in tariffs:
                days, price = tariff[1], tariff[3]
                if days == 'Навсегда':
                    kb.add(InlineKeyboardButton(f'{days} / {price} руб.', callback_data=f'{days}/{price}'))
                else:
                    kb.add(InlineKeyboardButton(f'{days} дней / {price} руб.', callback_data=f'{days}/{price}'))
            kb.add(InlineKeyboardButton('Изменить описание для всех тарифов канала',
                                        callback_data='change_description'))
            kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
            await call.message.answer(f'Тариф: {pick}\nОписание: {description}', reply_markup=kb)
            await Dialog.edit_tariff_start_state.set()


@dp.callback_query_handler(state=Dialog.edit_tariff_start_state)
async def edit_tariff_start_set(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        data = await state.get_data()
        tariff = data.get('choice_tariff')
        description = data.get('choice_description')
        if tariff == 'Всё вместе навсегда':
            kb = ReplyKeyboardMarkup()
            kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
            if pick == 'change_description':
                await call.message.answer(f'Тариф: {tariff}\nАктуальное описание: {description}\n\n'
                                          f'Введите новое описание:', reply_markup=kb)
                await state.update_data(choice_method='description')
                await Dialog.edit_tariff_start_state.set()
        else:
            if pick == 'change_description':
                kb = ReplyKeyboardMarkup()
                kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
                await call.message.answer(f'Тариф: {tariff}\nАктуальное описание: {description}\n\n'
                                          f'Введите новое описание:', reply_markup=kb)
                await state.update_data(choice_method='description')
                await Dialog.edit_tariff_start_state.set()
            else:
                kb = ReplyKeyboardMarkup()
                kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
                days, price = pick.split('/')
                await state.update_data(choice_days=days)
                await state.update_data(choice_price=price)
                if days == 'Навсегда':
                    await call.message.answer(f'Тариф: {tariff}\nЦена: {price} руб.\nСрок: {days}\nОписание: '
                                              f'{description}\n\nВведите новую цену для тарифа:', reply_markup=kb)
                else:
                    await call.message.answer(f'Тариф: {tariff}\nЦена: {price} руб.\nСрок: {days} дней\n'
                                              f'Описание: {description}\n\nВведите новую цену для тарифа:',
                                              reply_markup=kb)
                await state.update_data(choice_method='price')
                await Dialog.edit_tariff_start_state.set()


@dp.message_handler(state=Dialog.edit_tariff_start_state)
async def edit_tariff_start_set(msg: Message, state=FSMContext):
    await msg.delete()

    pick = msg.text
    if pick.lower() == 'меню↩️':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        data = await state.get_data()
        choice = data.get('choice_method')
        tariff = data.get('choice_tariff')
        description = data.get('choice_description')
        price = data.get('choice_price')
        days = data.get('choice_days')
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
        kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
        if choice == 'description':
            if tariff == 'Всё вместе навсегда':
                await msg.answer(f'Тариф: {tariff}\nЦена: {price}\nОписание: {pick}', reply_markup=kb)
            else:
                await msg.answer(f'Тариф: {tariff}\nОписание: {pick}', reply_markup=kb)
                await state.update_data(choice_description=pick)
        elif choice == 'price':
            if tariff == 'Всё вместе навсегда':
                await msg.answer(f'Тариф: {tariff}\nЦена: {pick}\nОписание: {description}', reply_markup=kb)
            else:
                await msg.answer(f'Тариф: {tariff}\nЦена: {pick}\nСрок: {days}\nОписание: {description}', reply_markup=kb)
            await state.update_data(choice_price=pick)
        else:
            await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()


@dp.callback_query_handler(state=Dialog.edit_tariff_ans_state)
async def edit_tariff_ans_set(call: CallbackQuery, state=FSMContext):
    if call.data == 'yes':
        data = await state.get_data()
        choice = data.get('choice_method')
        tariff = [data.get('choice_tariff'), data.get('choice_days'), data.get('choice_description'),
                  data.get('choice_price')]
        if choice == 'description':
            try:
                cur.execute(f'UPDATE Tariffs SET description="{tariff[2]}" WHERE name="{tariff[0]}" AND '
                            f'days="{tariff[1]}" AND price="{tariff[3]}" AND work=1')
            except:
                await call.message.answer('Произошла неизвестная ошибка\nИзменения отменены', reply_markup=menu_kb)
            else:
                await call.message.answer(f'Описание группы {tariff[0]} успешно изменена на: {tariff[2]}',
                                          reply_markup=menu_kb)
        elif choice == 'price':
            try:
                cur.execute(f'UPDATE Tariffs SET price="{tariff[3]}" WHERE name="{tariff[0]}" AND '
                            f'days="{tariff[1]}" AND description="{tariff[2]}" AND work=1')
            except:
                await call.message.answer('Произошла неизвестная ошибка\nИзменения отменены', reply_markup=menu_kb)
            else:
                await call.message.answer(f'Цена тарифа {tariff[0]}/{tariff[1]} успешно изменена на {tariff[3]} руб.',
                                          reply_markup=menu_kb)

        else:
            await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()
    else:
        await call.message.answer('Вы вернулись в меню\nИзменения отменены', reply_markup=menu_kb)
        await state.finish()


@dp.callback_query_handler(text='edit_forever')
async def edit_forever_func(call: CallbackQuery):
    cur.execute('SELECT * FROM Tariffs WHERE name="Всё вместе навсегда"')
    res = cur.fetchone()
    if res:
        price = res[3]
    else:
        price = 'Отсутствует'
    kb = ReplyKeyboardMarkup()
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    await call.message.answer(f'Цена для тарифа "Всё вместе навсегда": {price}\n'
                              f'Напишите новую цену', reply_markup=kb)
    await Dialog.edit_forever_state.set()


@dp.message_handler(state=Dialog.edit_forever_state)
async def edit_forever_set(msg: Message, state=FSMContext):
    price = msg.text.lower()
    if price == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        if price.isdigit():
            cur.execute('SELECT * FROM Tariffs WHERE name="Всё вместе навсегда"')
            res = cur.fetchone()
            if res:
                try:
                    cur.execute(f'UPDATE Tariffs SET price={price} WHERE name="Всё вместе навсегда"')
                    con.commit()
                except:
                    await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
                    await state.finish()
                else:
                    await msg.answer('Цена успешно изменена', reply_markup=menu_kb)
                    await state.finish()
            else:
                try:
                    cur.execute('INSERT INTO Tariffs VALUES (?, ?, ?, ?, ?)', ('Всё вместе навсегда', 'Навсегда',
                                                                               '', str(price), 1))
                    con.commit()
                except:
                    await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
                    await state.finish()
                else:
                    await msg.answer('Цена успешно изменена', reply_markup=menu_kb)
                    await state.finish()
        else:
            await msg.answer('Это не похоже на число', reply_markup=menu_kb)
            await state.finish()


@dp.callback_query_handler(text='new_phrase')
async def new_phrase_func(call: CallbackQuery):
    await call.message.delete()

    kb = ReplyKeyboardMarkup()
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    cur.execute('SELECT * FROM Settings')
    phrase = cur.fetchone()[2]
    await call.message.answer(f'Актуальная готовая фраза: {phrase} "Тариф"\nНапишите новую фразу', reply_markup=kb)
    await Dialog.edit_phrase_state.set()


@dp.message_handler(state=Dialog.edit_phrase_state)
async def edit_phrase_set(msg: Message, state=FSMContext):
    phrase = msg.text
    if phrase.lower() == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        try:
            cur.execute(f'UPDATE Settings SET phrase_tariff="{phrase}"')
            con.commit()
        except:
            await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
        else:
            await msg.answer(f'Новая фраза: {phrase} "Тариф"', reply_markup=menu_kb)
            await state.finish()


@dp.message_handler(state=Dialog.edit_phrase_state)
async def edit_phrase_set(msg: Message, state=FSMContext):
    phrase = msg.text
    if phrase.lower() == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        try:
            cur.execute(f'UPDATE Settings SET phrase_tariff={phrase}')
            con.commit()
        except:
            await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
        else:
            await msg.answer(f'Новая фраза: {phrase} "Тариф"', reply_markup=menu_kb)
            await state.finish()


@dp.callback_query_handler(text='stat')
async def stat_func(call: CallbackQuery):
    cur.execute('SELECT * FROM Users WHERE admin=0')
    users = cur.fetchall()
    try:
        bought = sum(tuple(map(lambda x: x[3], filter(lambda x: x[3], users))))
    except:
        bought = 'Ошибка'
    else:
        cur.execute('SELECT * FROM Subs')
        subs = cur.fetchall()
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton('Продажи', callback_data='show_purchases_list'))
        kb.add(InlineKeyboardButton('Пользователи', callback_data='show_users_list'))
        kb.add(InlineKeyboardButton('Подписчики', callback_data='show_subs_list'))
        kb.add(InlineKeyboardButton('Админы', callback_data='show_admins_list'))
        try:
            cur.execute('SELECT * FROM Purchase WHERE accept=1 and admin=0')
            purchases = list(map(lambda x: (int(x[5]), x[1]), cur.fetchall()))
            all = sum(tuple(map(lambda x: x[0], purchases)))
            week, in_month = 0, 0
            day, month, year = tuple(map(int, get_date().split('/')))
            now = datetime(year, month, day)
            for purchase in purchases:
                day1, month1, year1 = tuple(map(int, purchase[1].split('/')))
                date = datetime(year1, month1, day1)
                delta = str(now - date)
                if delta == '0:00:00':
                    delta = 0
                else:
                    delta = int(delta.split(',')[0].split()[0])
                if delta <= 7:
                    week += purchase[0]
                if delta <= 32:
                    in_month += purchase[0]
            all = sum(tuple(map(lambda x: int(x[0]), purchases)))
            today = sum(tuple(map(lambda x: int(x[0]), tuple(filter(lambda x: x[1] == get_date(), purchases)))))
        except:
            await call.message.answer(f'Переходов в бота: {len(users)}\nПокупок: {bought}\nАктивных подписок: {len(subs)}',
                                      reply_markup=kb)
        else:
            await call.message.answer(f'Переходов в бота: {len(users)}\nПокупок: {bought}\nАктивных подписок: {len(subs)}\n'
                                      f'\n\nЗаработано:\nВсего: {all} руб.\nЗа сегодня: {today} руб.\nЗа неделю: {week} руб.\n'
                                      f'За месяц: {in_month} руб.', reply_markup=kb)


@dp.callback_query_handler(text='show_purchases_list')
async def show_users_list_func(call: CallbackQuery):
    try:
        name = 'show_purchases_list.xlsx'
        workbook = Workbook(name)
        worksheet = workbook.add_worksheet()
        cur.execute('SELECT * FROM Purchase WHERE admin=0')
        cols = cur.description
        bd = cur.fetchall()
        for col in range(len(cols)):
            worksheet.write(0, col, str(cols[col][0]))
        for row in range(len(bd)):
            for col in range(len(bd[row])):
                worksheet.write(row + 1, col, str(bd[row][col]))
        workbook.close()
    except:
        await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
    else:
        await call.message.reply_document(open(name, 'rb'))


@dp.callback_query_handler(text='show_users_list')
async def show_users_list_func(call: CallbackQuery):
    try:
        name = 'show_users_list.xlsx'
        workbook = Workbook(name)
        worksheet = workbook.add_worksheet()
        cur.execute('SELECT * FROM Users')
        cols = cur.description
        bd = cur.fetchall()
        for col in range(len(cols)):
            worksheet.write(0, col, str(cols[col][0]))
        for row in range(len(bd)):
            for col in range(len(bd[row])):
                worksheet.write(row + 1, col, str(bd[row][col]))
        workbook.close()
    except:
        await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
    else:
        await call.message.reply_document(open(name, 'rb'))


@dp.callback_query_handler(text='show_subs_list')
async def show_subs_list_func(call: CallbackQuery):
    try:
        name = 'show_subs_list.xlsx'
        workbook = Workbook(name)
        worksheet = workbook.add_worksheet()
        cur.execute('SELECT * FROM Subs')
        cols = cur.description
        bd = cur.fetchall()
        for col in range(len(cols)):
            worksheet.write(0, col, str(cols[col][0]))
        for row in range(len(bd)):
            for col in range(len(bd[row])):
                worksheet.write(row + 1, col, str(bd[row][col]))
        workbook.close()
    except:
        await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
    else:
        await call.message.reply_document(open(name, 'rb'))


@dp.callback_query_handler(text='show_admins_list')
async def show_subs_list_func(call: CallbackQuery):
    try:
        name = 'show_admins_list.xlsx'
        workbook = Workbook(name)
        worksheet = workbook.add_worksheet()
        cur.execute('SELECT * FROM Admins')
        cols = cur.description
        bd = cur.fetchall()
        for col in range(len(cols)):
            worksheet.write(0, col, str(cols[col][0]))
        for row in range(len(bd)):
            for col in range(len(bd[row])):
                worksheet.write(row + 1, col, str(bd[row][col]))
        workbook.close()
    except:
        await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
    else:
        await call.message.reply_document(open(name, 'rb'))


@dp.callback_query_handler(text='search')
async def search_func(call: CallbackQuery):
    await call.message.answer('Функция временно недоступна', reply_markup=menu_kb)
    pass


@dp.callback_query_handler(text='edit_helper')
async def search_func(call: CallbackQuery, state=FSMContext):
    cur.execute('SELECT * FROM Settings')
    helper = cur.fetchone()[3]
    await state.update_data(helper=helper)
    kb = ReplyKeyboardMarkup()
    kb.add(InlineKeyboardButton('Удалить контакт тех. поддержки', callback_data='delete_helper'))
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    if helper:
        await call.message.answer(f'Тех. поддержка: {helper}\nНапишите контакт новой тех. поддержки',
                                  reply_markup=kb)
    else:
        await call.message.answer('Напишите контакт новой тех. поддержки', reply_markup=kb)
    await Dialog.edit_helper_state.set()


@dp.message_handler(state=Dialog.edit_helper_state)
async def edit_helper_set(msg: Message, state=FSMContext):
    pick = msg.text
    if pick.lower() == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
    elif pick.lower() == 'удалить контакт тех. поддержки':
        try:
            cur.execute(f'UPDATE Settings SET helper="" WHERE rowid=1')
        except:
            await msg.answer('Не удалось изменить контакт тех. поддержки', reply_markup=menu_kb)
        else:
            await msg.answer('Контакт тех. поддержки удалён', reply_markup=menu_kb)
    else:
        try:
            cur.execute(f'UPDATE Settings SET helper="{pick}" WHERE rowid=1')
        except:
            await msg.answer('Не удалось изменить контакт тех. поддержки', reply_markup=menu_kb)
        else:
            data = await state.get_data()
            await msg.answer(f'{data.get("helper")} успешно изменён на {pick}', reply_markup=menu_kb)
    await state.finish()


@dp.callback_query_handler(text='edit_deposit')
async def edit_helper_set(call: CallbackQuery, state=FSMContext):
    cur.execute('SELECT * FROM Settings')
    mn = cur.fetchone()[4]
    cur.execute('SELECT * FROM Payment_methods')
    payments = len(cur.fetchall())
    kb = InlineKeyboardMarkup()
    kb.add(InlineKeyboardButton('Минимальная сумма для пополнения', callback_data='edit_min_deposit'))
    kb.add(InlineKeyboardButton('Изменить способ оплаты', callback_data='edit_payment_method'))
    kb.add(InlineKeyboardButton('Добавить способ оплаты', callback_data='add_payment_method'))
    await call.message.answer(f'Минимальная сумма для пополнения: {mn}\nСпособов оплаты: {payments}\n'
                              f'Выберите действие', reply_markup=kb)


@dp.callback_query_handler(text='edit_min_deposit')
async def edit_min_deposit_func(call: CallbackQuery):
    kb = ReplyKeyboardMarkup()
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    await call.message.answer('Напишите минимальную сумму (целое число)', reply_markup=kb)
    await Dialog.edit_min_deposit_state.set()


@dp.message_handler(state=Dialog.edit_min_deposit_state)
async def edit_min_deposit_set(msg: Message, state=FSMContext):
    pick = msg.text.lower()
    if pick == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
    elif pick.isdigit():
        pick = int(pick)
        if pick >= 0:
            try:
                cur.execute(f'UPDATE Settings SET min_deposit={pick}')
                con.commit()
            except:
                await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            else:
                await msg.answer('Минимальная сумма изменена', reply_markup=menu_kb)
        else:
            await msg.answer('Отрицательное число не может быть минимальной суммой для пополнения',
                             reply_markup=menu_kb)
            await state.finish()
    else:
        await msg.answer('Вы неверно ввели сумму', reply_markup=menu_kb)
    await state.finish()


@dp.callback_query_handler(text='edit_payment_method')
async def edit_min_deposit_func(call: CallbackQuery):
    cur.execute('SELECT * FROM Payment_methods')
    methods = sorted(list(set(map(lambda x: x[0], cur.fetchall()))))
    kb = InlineKeyboardMarkup()
    for method in methods:
        kb.add(InlineKeyboardButton(method, callback_data=method))
    kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
    await call.message.answer('Изменение способов оплаты', reply_markup=ReplyKeyboardRemove())
    await call.message.answer('Выберите способ оплаты', reply_markup=kb)
    await Dialog.edit_payment_method_state.set()


@dp.callback_query_handler(state=Dialog.edit_payment_method_state)
async def edit_payment_method_set(call: CallbackQuery, state=FSMContext):
    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        await state.update_data(choice_pay_method=pick)
        cur.execute(f'SELECT * FROM Payment_methods WHERE name="{pick}" AND work=1')
        method = cur.fetchone()
        await state.update_data(choice_pay_method=method[0])
        await state.update_data(old_requisites=method[3])
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton('Изменить название', callback_data='edit_name'))
        kb.add(InlineKeyboardButton('Изменить реквизиты', callback_data='edit_requisites'))
        kb.add(InlineKeyboardButton('Удалить', callback_data='delete_method'))
        kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
        await call.message.answer(f'{pick}\nВыберите действие', reply_markup=kb)
        await Dialog.choice_edit_payment_method_state.set()


@dp.callback_query_handler(state=Dialog.choice_edit_payment_method_state)
async def choice_edit_payment_method_set(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        data = await state.get_data()
        method = data.get('choice_pay_method')
        kb = ReplyKeyboardMarkup()
        kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
        if pick == 'edit_name':
            await call.message.answer(f'Напишите новое название для {method}', reply_markup=kb)
            await Dialog.edit_name_state.set()
        elif pick == 'edit_requisites':
            await call.message.answer(f'Напишите новые реквизиты для {method}', reply_markup=kb)
            await Dialog.edit_requisites_state.set()
        elif pick == 'delete_method':
            kb = InlineKeyboardMarkup()
            kb.add(InlineKeyboardButton('Удалить', callback_data='yes'))
            kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
            await call.message.answer(f'Удалить "{method}"\nВы уверены?', reply_markup=kb)
            await Dialog.delete_payment_method_state.set()


@dp.message_handler(state=Dialog.edit_name_state)
async def edit_name_set(msg: Message, state=FSMContext):
    await msg.delete()

    name = msg.text
    if name.lower() == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        data = await state.get_data()
        method = data.get('choice_pay_method')
        name = name.capitalize()
        if method == name:
            await msg.answer('Такой способ оплаты уже существует', reply_markup=menu_kb)
            await state.finish()
        else:
            await state.update_data(new_method=name)
            kb = InlineKeyboardMarkup()
            kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
            kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
            await msg.answer(f'Заменить {method} на {name}?', reply_markup=kb)
            await Dialog.edit_name_ans_state.set()


@dp.callback_query_handler(state=Dialog.edit_name_ans_state)
async def edit_name_ans_set(call: CallbackQuery, state=FSMContext):
    if call.data == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        try:
            data = await state.get_data()
            new, old = data.get('new_method'), data.get('choice_pay_method')
            cur.execute(f'UPDATE Payment_methods SET name="{new}" WHERE name="{old}" AND work=1')
            con.commit()
        except:
            await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()
        else:
            await call.message.answer(f'{old} успешно изменён на {new}', reply_markup=menu_kb)
            await state.finish()


@dp.callback_query_handler(state=Dialog.edit_requisites_state)
async def edit_requisites_set(call: CallbackQuery, state=FSMContext):
    data = await state.get_data()
    method = data.get('choice_pay_method')
    kb = ReplyKeyboardMarkup()
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    await call.message.answer(f'{method}\nВведите новые реквизиты', reply_markup=kb)
    await Dialog.choice_edit_requisites_state.set()


@dp.message_handler(state=Dialog.choice_edit_requisites_state)
async def choice_edit_requisites_set(msg: Message, state=FSMContext):
    await msg.delete()

    pick = msg.text.lower()
    if pick == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        data = await state.get_data()
        old = data.get('old_requisites')
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
        kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
        await msg.answer(f'{data.get("choice_pay_method")}\nЗаменить {old} на {pick}?', reply_markup=kb)
        await Dialog.edit_requisites_ans_state.set()


@dp.callback_query_handler(state=Dialog.edit_requisites_ans_state)
async def edit_requisites_ans_set(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    if call.data == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        try:
            data = await state.get_data()
            old, new, method = data.get('old_requisites'), data.get('new_requisites'), data.get("choice_pay_method")
            cur.execute(f'UPDATE Payment_methods SET number="{new}" WHERE name="{method}" AND number="{old}"')
            con.commit()
        except:
            await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()
        else:
            await call.message.answer(f'{method}\n{old} успешно заменён на {new}', reply_markup=menu_kb)
            await state.finish()



@dp.callback_query_handler(text='add_payment_method')
async def edit_min_deposit_func(call: CallbackQuery):
    kb = ReplyKeyboardMarkup()
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    await call.message.answer('Напишите новый способ оплаты в виде:\nНазвание/реквизиты', reply_markup=kb)
    await Dialog.add_payment_method_state.set()


@dp.message_handler(state=Dialog.add_payment_method_state)
async def add_payment_method_set(msg: Message, state=FSMContext):
    pick = msg.text
    if pick.lower() == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        pick = pick.split('/')
        if len(pick) != 2:
            await msg.answer('В введённых вами данных ошибка', reply_markup=menu_kb)
            await state.finish()
        else:
            cur.execute(f'SELECT * FROM Payment_methods WHERE name="{pick[0]}"')
            if cur.fetchone():
                await msg.answer('Способ с таким названием уже существует', reply_markup=menu_kb)
                await state.finish()
            else:
                await state.update_data(choice_pay_method=pick)
                kb = InlineKeyboardMarkup()
                kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
                kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
                await msg.answer(f'Название: {pick[0]}\nРеквизиты: {pick[1]}\n\nПодтвердить?', reply_markup=kb)
                await Dialog.add_payment_method_ans_state.set()


@dp.callback_query_handler(state=Dialog.add_payment_method_ans_state)
async def add_payment_method_ans_set(call: CallbackQuery, state=FSMContext):
    if call.data == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        try:
            data = await state.get_data()
            new = data.get('choice_pay_method')
            cur.execute('INSERT INTO Payment_methods VALUES (?, ?, ?, ?, ?, ?)', (new[0], '', '', new[1], 0, 1))
            con.commit()
        except:
            await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()
        else:
            await call.message.answer(f'Способ оплаты "{new[0]}" успешно добавлен', reply_markup=menu_kb)
            await state.finish()


@dp.callback_query_handler(state=Dialog.delete_payment_method_state)
async def delete_payment_method_set(call: CallbackQuery, state=FSMContext):
    if call.data == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_kb)
        await state.finish()
    else:
        try:
            data = await state.get_data()
            method = data.get('choice_pay_method')
            cur.execute(f'DELETE FROM Payment_methods WHERE name="{method}" AND work=1')
            con.commit()
        except:
            await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
            await state.finish()
        else:
            await call.message.answer(f'"{method}" успешно удалён', reply_markup=menu_kb)
            await state.finish()


@dp.message_handler(text='🥰Тарифы🥰')
async def tariffs_btn(msg: Message):
    cur.execute(f'SELECT * FROM Tariffs WHERE work=1 AND name != "Всё вместе навсегда"')
    tariffs = sorted(tuple(set(map(lambda x: x[0], cur.fetchall()))))
    if tariffs:
        kb = InlineKeyboardMarkup()
        for i in range(len(tariffs)):
            tariff = tariffs[i]
            kb.add(InlineKeyboardButton(tariff, callback_data=tariff))
        cur.execute('SELECT * FROM Settings')
        if cur.fetchone()[1]:
            try:
                cur.execute('SELECT * FROM Tariffs WHERE work=1 AND days="Навсегда"')
                tariff = cur.fetchone()[0]
                # kb.add(InlineKeyboardButton('Всё вместе навсегда', callback_data="Всё вместе навсегда"))
            except:
                pass
        kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
        await msg.answer('Добро пожаловать в магазин!', reply_markup=ReplyKeyboardRemove())
        await msg.answer('Выберите желаемый тарифный план:', reply_markup=kb)
        await (Dialog.choice_time_state.set())
    else:
        await msg.answer('Доступные тарифы временно отсутствуют', reply_markup=menu_user)


@dp.callback_query_handler(state=Dialog.choice_time_state)
async def choice_time_set(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_user)
        await state.finish()
    else:
        await state.update_data(deposit_buy='buy')
        await state.update_data(choice_tariff=pick)
        await state.update_data(user_id=call.from_user.id)
        data = await state.get_data()
        if pick == "Всё вместе навсегда":
            await state.update_data(choice_days='Навсегда')
            data = await state.get_data()
            cur.execute('SELECT * FROM Payment_methods WHERE work=1')
            methods = cur.fetchall()
            kb = InlineKeyboardMarkup()
            for i in range(len(methods)):
                method = methods[i][0]
                kb.add(InlineKeyboardButton(method, callback_data=method))
            kb.add(InlineKeyboardButton('Баланс', callback_data='Баланс'))
            kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
            cur.execute(f'SELECT * FROM Tariffs WHERE name="{data.get("choice_tariff")}" AND '
                        f'days="{data.get("choice_days")}"')
            tariff = cur.fetchone()
            try:
                await state.update_data(choice_price=tariff[3])
            except:
                await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
                await state.finish()
            else:
                cur.execute('SELECT * FROM Tariffs WHERE work=1 AND name != "Всё вместе навсегда"')
                tariffs = '\n'.join(sorted(tuple(filter(lambda x: x.strip() != 'all', set(map(lambda x: x[0], cur.fetchall()))))))
                await call.message.answer(f'Тариф: Всё вместе навсегда\nСрок: Навсегда\nЦена: {tariff[3]}\n'
                                          f'Вы получите приглашение в каналы/чаты 👇\n{tariffs}\n\nВыберите метод оплаты 👇',
                                          reply_markup=kb)
                await Dialog.accept_pay_state.set()
        else:
            cur.execute(f'SELECT * FROM Tariffs WHERE name="{pick}"')
            times = cur.fetchall()
            if times:
                kb = InlineKeyboardMarkup()
                for i in range(len(times)):
                    time = times[i][1]
                    if time == 'Навсегда':
                        kb.add(InlineKeyboardButton(f'{time}', callback_data=time))
                    else:
                        kb.add(InlineKeyboardButton(f'{time} дней', callback_data=time))
                kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
                await call.message.answer('Выберите тариф или категорию из списка ниже 👇', reply_markup=kb)
                await state.update_data(deposit_buy='buy')
                await Dialog.choice_pay_method_state.set()
            else:
                await call.message.answer('Данный тариф временно недоступен', reply_markup=menu_user)
                await state.finish()


@dp.callback_query_handler(state=Dialog.choice_pay_method_state)
async def choice_pay_method_set(call: CallbackQuery, state=FSMContext):
    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_user)
        await state.finish()
    else:
        data = await state.get_data()
        if data.get('deposit_buy') == 'buy':
            await state.update_data(choice_days=pick)
            data = await state.get_data()
            cur.execute('SELECT * FROM Payment_methods WHERE work=1')
            methods = cur.fetchall()
            kb = InlineKeyboardMarkup()
            for i in range(len(methods)):
                method = methods[i][0]
                kb.add(InlineKeyboardButton(method, callback_data=method))
            kb.add(InlineKeyboardButton('Баланс', callback_data='Баланс'))
            kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
            cur.execute(f'SELECT * FROM Tariffs WHERE name="{data.get("choice_tariff")}" AND '
                        f'days = "{data.get("choice_days")}" AND work=1')
            tariff = cur.fetchone()
            await state.update_data(choice_price=tariff[3])
            if pick == 'Навсегда':
                await call.message.answer(f'Канал: {data.get("choice_tariff")}\nСрок: {data.get("choice_days")}\n'
                                          f'Описание: {tariff[2]}\nЦена: {tariff[3]} ₽\n\nВыберите метод оплаты 👇',
                                          reply_markup=kb)
                await Dialog.accept_pay_state.set()
            else:
                await call.message.answer(f'Канал: {data.get("choice_tariff")}\nСрок: {data.get("choice_days")} дней\n'
                                        f'Описание: {tariff[2]}\nЦена: {tariff[3]} ₽\n\nВыберите метод оплаты 👇',
                                        reply_markup=kb)
                await Dialog.accept_pay_state.set()
        elif data.get('deposit_buy') == 'deposit':
            await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
            await state.finish()


@dp.callback_query_handler(state=Dialog.accept_pay_state)
async def accept_pay_set(call: CallbackQuery, state=FSMContext):
    pick = call.data
    if pick == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_user)
        await state.finish()
    else:
        data = await state.get_data()
        price = int(data.get('choice_price'))
        await state.update_data(choice_method=pick)
        if pick == 'Баланс':
            cur.execute(f'SELECT * FROM Users WHERE user_id={call.from_user.id}')
            balance = int(cur.fetchone()[4])
            await state.update_data(balance=balance)
            if balance < price:
                kb = InlineKeyboardMarkup()
                kb.add(InlineKeyboardButton('Пополнить баланс', callback_data='deposit'))
                await call.message.answer(f'Цена тарифа: {price}\nВаш баланс: {balance}', reply_markup=menu_user)
                await call.message.answer('Недостаточно средств на балансе', reply_markup=kb)
                await state.finish()
            else:
                kb = InlineKeyboardMarkup()
                kb.add(InlineKeyboardButton('Да', callback_data='yes'))
                kb.add(InlineKeyboardButton('Нет', callback_data='to_menu'))
                cur.execute(f'SELECT * FROM Tariffs WHERE name="{data.get("choice_tariff")}" AND '
                            f'days="{data.get("choice_days")}" AND work=1')
                tariff = cur.fetchone()
                if tariff:
                    await state.update_data(choice_tariff=tariff)
                    await call.message.answer(f'Канал: {tariff[0]}\nДней: {tariff[1]}\nОписание: {tariff[2]}'
                                              f'Цена: {price} руб.\nСпособ оплаты: Баланс\n\nВы уверены?',
                                              reply_markup=kb)
                    await Dialog.balance_pay_check_state.set()
                else:
                    await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
                    await state.finish()
        else:
            cur.execute(f'SELECT * FROM Payment_methods WHERE name="{pick}" AND work=1')
            method = cur.fetchone()
            if method:
                deposit_buy = data.get('deposit_buy')
                if method[4]:
                    pass
                else:
                    kb = InlineKeyboardMarkup()
                    kb.add(InlineKeyboardButton('Я оплатил', callback_data='yes'))
                    kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
                    await call.message.answer(f'Способ оплаты: {method[0]}\n'
                                              f'Сумма к оплате: {data.get("choice_price")}\nПеревод: {method[3]}',
                                              reply_markup=kb)
                    await Dialog.pay_fiat_state.set()
            else:
                await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
                await state.finish()


@dp.callback_query_handler(state=Dialog.pay_fiat_state)
async def check_pay_fiat_set(call: CallbackQuery, state=FSMContext):
    if call.data == 'to_menu':
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_user)
        await state.finish()
    else:
        kb = ReplyKeyboardMarkup()
        kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
        await call.message.answer('💰 Оплатили?\n\nОтправьте боту квитанцию об оплате: скриншот или фото.\n'
                                  'На квитанции должны быть четко видны: дата, время и сумма платежа.\n'
                                  'Для отмены нажмите кнопку "Отмена"', reply_markup=kb)
        await Dialog.check_pay_fiat_state.set()


@dp.message_handler(content_types=['photo', 'text'], state=Dialog.check_pay_fiat_state)
async def check_pay_fiat_set(msg: Message, state=FSMContext):
    if msg.photo:
        photo = msg.photo[0].file_id
        date = get_date()
        data = await state.get_data()
        user_name = msg.from_user.full_name
        user_id = msg.from_user.id
        price = data.get("choice_price")
        method = data.get("choice_method")
        cur.execute('SELECT * FROM Purchase')
        number = len(cur.fetchall())
        deposit_buy = data.get('deposit_buy')
        cur.execute('SELECT * FROM Admins WHERE payments=1')
        admins = tuple(map(lambda x: x[0], cur.fetchall()))
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton('✅Подтвердить', callback_data='accept_purchase'))
        kb.add(InlineKeyboardButton('❌Спам', callback_data='cancel_purchase'))
        if data.get('choice_promo') == None:
            await state.update_data(choice_promo='Отсутствует')
            data = await state.get_data()
        if deposit_buy == 'buy':
            tariff = f'{data.get("choice_tariff")}/{data.get("choice_days")}'
        elif deposit_buy == 'deposit':
            tariff = deposit_buy
        admin = 0
        if msg.from_user.id in admins:
            admin = 1
        cur.execute('INSERT INTO Purchase VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', (user_id, date, photo, tariff, method,
                                                                                price, number, False, False, admin))
        con.commit()
        for admin in admins:
            await bot.send_photo(chat_id=admin, photo=photo, caption=f'💰Подтвердите покупку\n\nПользователь: '
                                                                         f'{user_name}\n'
                                                                         f'user id: {user_id}\n'
                                                                         f'Тариф: {tariff}\n'
                                                                         f'Промокод: {data.get("choice_promo")}\n'
                                                                         f'Сумма к оплате: {price}\n'
                                                                         f'Платёжная система: {method}\n'
                                                                         f'Номер платежа: {number}', reply_markup=kb)
        await state.finish()
    elif msg.text.lower() == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_user)
        await state.finish()
    else:
        await msg.answer('Сообщенине не является фотографией (напишите "Отмена" для отмены)')


@dp.callback_query_handler(text='accept_purchase')
async def accept_purchase_func(call: CallbackQuery):
    photo = dict(call)['message']['photo'][0]['file_id']
    await call.message.answer(f'Введите номер платежа без каких-либо знаков препинания')
    await Dialog.accept_purchase_state.set()


@dp.message_handler(state=Dialog.accept_purchase_state)
async def accept_purchase_set(msg: Message, state=FSMContext):
    number = msg.text.strip()
    try:
        cur.execute(f'SELECT * FROM Purchase WHERE number={number}')
        data = cur.fetchone()
    except:
        await msg.answer('Неверно введён номер платежа. Нажмите '
                         '"✅Подтвердить" и введите номер платежа заново', reply_markup=menu_kb)
        await state.finish()
    else:
        if not data[7]:
            number = int(number)
            cur.execute('SELECT * FROM Purchase')
            all = len(cur.fetchall())
            if number > all:
                await msg.answer('Неверно введён номер платежа. Нажмите '
                                 '"✅Подтвердить" и введите номер платежа заново', reply_markup=menu_kb)
                await state.finish()
            elif data[7]:
                await msg.answer('Неверно введён номер платежа. Нажмите '
                                 '"✅Подтвердить" и введите номер платежа заново', reply_markup=menu_kb)
                await state.finish()
            else:
                tariff = data[3].split('/')[0]
                if tariff == 'deposit':
                    try:
                        cur.execute(f'SELECT * FROM Purchase WHERE number={number} AND accept=0')
                        data = cur.fetchone()
                    except:
                        await msg.answer(f'Произошла неизвестная ошибка по платежу # {number}')
                    else:
                        if data:
                            try:
                                cur.execute(f'SELECT * FROM Users WHERE user_id={data[0]}')
                                user = cur.fetchone()
                                balance = user[4]
                                cur.execute(f'UPDATE Users SET balance={int(balance) + int(data[5])} WHERE user_id='
                                            f'{user[0]}')
                                con.commit()
                            except:
                                await msg.answer(f'Произошла неизвестная ошибка по платежу # {number}')
                            else:
                                await msg.answer(f'Платёж # {number} подтверждён')
                                await bot.send_message(chat_id=data[0], text=f'Платёж #{number} на сумму {data[5]} руб. '
                                                                             f'подтверджён', reply_markup=menu_user)
                                cur.execute(f'UPDATE Purchase SET sent=1 WHERE number={number}')
                                cur.execute(f'UPDATE Purchase SET accept=1 WHERE number={number}')
                                con.commit()
                            await state.finish()
                        else:
                            await msg.answer(f'Платёж # {number} уже подтверждён')
                else:
                    link, links = None, None
                    try:
                        if tariff == 'Всё вместе навсегда':
                            cur.execute('SELECT * FROM Tariffs_links')
                            links = list(map(lambda x: x[1], cur.fetchall()))
                            for i in range(len(links)):
                                link = await bot.export_chat_invite_link(chat_id=links[i])
                                links[i] = link
                            links = '\n'.join(links)
                            link = None
                        else:
                            cur.execute(f'SELECT * FROM Tariffs_links WHERE name="{data[3].split("/")[0]}"')
                            link = tuple(cur.fetchone())[1]
                            link = await bot.export_chat_invite_link(chat_id=str(link).strip())
                    except:
                        await msg.answer(f'❗Внимание❗\nНе удалось отправить ссылку на вступление по платежу # {number}',
                                         reply_markup=menu_kb)
                        await bot.send_message(chat_id=data[0],
                                               text=f'Платёж #{number} подтверджён\nНе удалось отправить ссылку на канал. '
                                                    f'Обратитесь в тех. поддержку',
                                               reply_markup=menu_user)
                    else:
                        await msg.answer(f'Ссылка на вступление в канал отправлена пользователю по платежу #{number} отправлена',
                                         reply_markup=menu_kb)
                        if link:
                            await bot.send_message(chat_id=data[0], text=f'Платёж #{number} подтверджён\nОдноразовая ссылка для вступления: '
                                                                         f'{link}', reply_markup=menu_user)
                        else:
                            await bot.send_message(chat_id=data[0],
                                                   text=f'Платёж #{number} подтверджён\nОдноразовые ссылки для вступления: '
                                                        f'{links}', reply_markup=menu_user)
                        cur.execute('INSERT INTO Subs VALUES (?, ?, ?, ?)', (data[0], data[3].split("/")[1], data[3].split()[0],
                                                                             get_date()))
                        cur.execute(f'UPDATE Purchase SET accept=1 WHERE number={number}')
                        cur.execute(f'UPDATE Users SET bought=1 WHERE user_id={data[0]}')
                        con.commit()
                    await state.finish()
        else:
            await msg.answer('Неверно введён номер платежа. Нажмите '
                             '"✅Подтвердить" и введите номер платежа заново', reply_markup=menu_kb)
            await state.finish()


@dp.callback_query_handler(text='cancel_purchase')
async def cancel_purchase_func(call: CallbackQuery):
    photo = dict(call)['message']['photo'][0]['file_id']
    await call.message.answer(f'Введите номер платежа без каких-либо знаков препинания')
    await Dialog.cancel_purchase_state.set()


@dp.message_handler(state=Dialog.cancel_purchase_state)
async def cancel_purchase_set(msg: Message, state=FSMContext):
    number = msg.text.strip()
    try:
        cur.execute(f'SELECT * FROM Purchase WHERE number={number} AND accept=0')
        data = cur.fetchone()
    except:
            await msg.answer('Неверно введён номер платежа. Нажмите "❌Спам" и введите номер платежа заново',
                             reply_markup=menu_kb)
    else:
        if data:
            await msg.answer(f'Платёж #{number} отменён', reply_markup=menu_kb)
            await bot.send_message(chat_id=data[0], text=f'Платёж #{number} не действителен. '
                                                     f'В случае ошибки обратитесь в тех. поддержку', reply_markup=menu_user)
        else:
            await msg.answer('Неверно введён номер платежа. Нажмите "❌Спам" и введите номер платежа заново',
                             reply_markup=menu_kb)
    finally:
        await state.finish()


@dp.callback_query_handler(state=Dialog.balance_pay_check_state)
async def balance_pay_check_set(call: CallbackQuery, state=FSMContext):
    await call.message.delete()

    if call.data == 'yes':
        try:
            data = await state.get_data()
            tariff = data.get('choice_tariff')
            name = tariff[0]
            if name == 'Всё вместе навсегда':
                cur.execute('SELECT * FROM Tariffs_links')
                links = list(map(lambda x: x[1], cur.fetchall()))
                for i in range(len(links)):
                    link = await bot.create_chat_invite_link(chat_id=links[i], member_limit=1)
                    links[i] = link['invite_link']
                links = '\n'.join(links)
                link = None
            else:
                cur.execute(f'SELECT * FROM Tariffs_links WHERE name="{name}"')
                link = tuple(cur.fetchone())[1]
                link = await bot.create_chat_invite_link(chat_id=str(link).strip(), member_limit=1)
                link = link['invite_link']
        except:
            await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
            await state.finish()
        else:
            user_id = call.from_user.id
            date = get_date()
            price = tariff[3]
            days = tariff[1]
            admin = 0
            if call.from_user.id in ADMINS:
                admin = 1
            try:
                if link:
                    await call.message.answer(f'Одноразовая ссылка для вступления: {link}',
                                              reply_markup=menu_user)
                elif links:
                    await call.message.answer(f'Одноразовые ссылки для вступления: {links}',
                                              reply_markup=menu_user)
            except:
                cur.execute('INSERT INTO Purchase VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', (user_id, date, '',
                                                                                 f'{name}/{tariff[1]}', 'Баланс',
                                                                                 price, 0, 0, 0, admin))
                await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
                await state.finish()
            else:
                if days == 'Навсегда':
                    days = 0
                cur.execute(f'UPDATE Users SET balance={int(data.get("balance")) - int(price)} WHERE '
                            f'user_id="{user_id}"')
                cur.execute('SELECT * FROM Purchase')
                length = len(cur.fetchall())
                cur.execute('INSERT INTO Purchase VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);', (user_id, date, None,
                                                                                            f'{name}/{tariff[1]}',
                                                                                            'Баланс', price, length,
                                                                                            1, 1, admin))
                cur.execute('INSERT INTO Subs VALUES (?, ?, ?, ?);', (user_id, days, name, date))
                con.commit()
                await state.finish()
    else:
        await call.message.answer('Вы вернулись в меню', reply_markup=menu_user)
        await state.finish()


@dp.message_handler(text='🌟Мой профиль🌟')
async def my_profile_func(msg: Message):
    cur.execute(f'SELECT * FROM Users WHERE user_id={msg.from_user.id}')
    user = cur.fetchone()
    kb = InlineKeyboardMarkup()
    kb.add(InlineKeyboardButton('Пополнить баланс', callback_data='deposit'))
    await msg.answer(f'id: {user[0]}\nДата регистрации: {user[2]}\nБаланс: {user[4]}', reply_markup=kb)


@dp.callback_query_handler(text='deposit')
async def deposit_sum_func(call: CallbackQuery, state=FSMContext):
    await state.update_data(deposit_buy='deposit')
    cur.execute('SELECT * FROM Settings')
    mn = cur.fetchone()[4]
    await state.update_data(choice_price=mn)
    kb = ReplyKeyboardMarkup()
    kb.add(InlineKeyboardButton('Отмена', callback_data='to_menu'))
    await call.message.answer('Пополнение баланса')
    await call.message.answer(f'Введите сумму для пополнения (от {mn})', reply_markup=kb)
    await Dialog.deposit_state.set()


@dp.message_handler(state=Dialog.deposit_state)
async def deposit_set(msg: Message, state=FSMContext):
    pick = msg.text.lower()
    if pick == 'отмена':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_user)
        await state.finish()
    elif pick.isdigit():
        pick = int(pick)
        data = await state.get_data()
        mn = data.get('choice_price')
        if not str(mn).isdigit():
            mn = 0
        if pick >= mn:
            await state.update_data(choice_price=pick)
            cur.execute('SELECT * FROM Payment_methods WHERE work=1')
            methods = cur.fetchall()
            kb = InlineKeyboardMarkup()
            for i in range(len(methods)):
                method = methods[i][0]
                kb.add(InlineKeyboardButton(method, callback_data=method))
            kb.add(InlineKeyboardButton('Меню↩️', callback_data='to_menu'))
            await msg.answer(f'Сумма: {pick}\nВыберите метод оплаты', reply_markup=kb)
            await Dialog.accept_pay_state.set()
        else:
            await msg.answer(f'Сумма меньше минимальной для пополнения (пополнение от {mn} руб.)', reply_markup=menu_kb)
            await state.finish()
    else:
        await msg.answer('Вы неверно ввели сумму', reply_markup=menu_user)
        await state.finish()


@dp.message_handler(text='📞Обратная связь📞')
async def callback_func(msg: Message):
    cur.execute('SELECT * FROM Settings')
    helper = cur.fetchone()[3]
    if helper:
        await msg.answer(f'Тех. поддержка: {helper}')
    else:
        await msg.answer('Тех. поддержка временно недоступна')


@dp.message_handler(text='🔥Предложить себя🔥')
async def offer_me(msg: Message, state=FSMContext):
    temp_kb = ReplyKeyboardMarkup()
    temp_kb.add(InlineKeyboardButton('Меню', callback_data='to_menu'))
    await msg.answer('Отправьте фотографию или видео, если хотите чтобы мы выложили ваш контент в наш канал',
                              reply_markup=temp_kb)
    await Dialog.offer_me_state.set()


@dp.message_handler(content_types=['text', 'photo', 'video'], state=Dialog.offer_me_state)
async def offer_me_set(msg: Message, state=FSMContext):
    if msg.text == 'Меню':
        await msg.answer('Вы вернулись в меню', reply_markup=menu_user)
        await state.finish()
    else:
        temp_kb = InlineKeyboardMarkup()
        temp_kb.add(InlineKeyboardButton('Подтвердить', callback_data='yes'))
        temp_kb.add(InlineKeyboardButton('Отмена', callback_data='no'))
        caption = ''
        if msg['caption']:
            caption = msg['caption']
        if msg.photo or msg.video:
            if msg.photo:
                await state.update_data(type='photo')
                media = msg.photo[0].file_id
                await state.update_data(media=media)
                await state.update_data(caption=caption)
                try:
                    await bot.send_photo(chat_id=msg.from_user.id, photo=media, caption=caption)
                except:
                    await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
                else:
                    await msg.answer('При нажатии "Подтвердить", вы соглашаетесь с тем, что мы выложим ваши материалы в наш '
                                     'публичный канал', reply_markup=temp_kb)
            else:
                await state.update_data(type='video')
                media = msg.video.file_id
                await state.update_data(media=media)
                await state.update_data(caption=caption)
                try:
                    await bot.send_video(chat_id=msg.from_user.id, video=media, caption=caption)
                except:
                    await msg.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
                else:
                    await msg.answer(
                        'При нажатии "Подтвердить", вы соглашаетесь с тем, что мы выложим ваши материалы в наш '
                        'публичный канал', reply_markup=temp_kb)
            await Dialog.ans_offer_me_state.set()
        else:
            await msg.answer('Сообщение не является медиафайлом', reply_markup=menu_user)
            await state.finish()


@dp.callback_query_handler(state=Dialog.ans_offer_me_state)
async def offer_me_photo_set(call: CallbackQuery, state: FSMContext):
    if call.data == 'yes':
        cur.execute('SELECT * FROM Media')
        medias = len(cur.fetchall()) + 1
        temp_kb = InlineKeyboardMarkup()
        temp_kb.add(InlineKeyboardButton('Подтвердить', callback_data=f'accept_offer_{medias}'))
        temp_kb.add(InlineKeyboardButton('Нет', callback_data=f'cancel_offer_{medias}'))
        try:
            cur.execute('SELECT * FROM Admins')
            admins = tuple(map(lambda x: x[0], cur.fetchall()))
            data = await state.get_data()
            media = data.get('media')
            type = data.get('type')
            caption = data.get('caption')
            if type == 'photo':
                for admin in admins:
                    await bot.send_photo(chat_id=admin, photo=media, caption=f'Предложение #{medias}\n{caption}',
                                         reply_markup=temp_kb)
            elif type == 'video':
                for admin in admins:
                    await bot.send_video(chat_id=admin, video=media, caption=f'Предложение #{medias}\n{caption}',
                                         reply_markup=temp_kb)
        except:
            await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_user)
        else:
            await call.message.answer('Предложение отправлено на проверку', reply_markup=menu_user)
            try:
                cur.execute('INSERT INTO Media VALUES (?, ?, ?, ?, ?, ?);', (call.from_user.id, media, caption,
                                                                                type, 0, medias))
            except:
                pass
            else:
                con.commit()
    await call.message.answer('Вы вернулись в меню', reply_markup=menu_user)
    await state.finish()


@dp.callback_query_handler(text_startswith='accept_offer_')
async def offer_me(call: CallbackQuery):
    number = str(call.data).split('accept_offer_')
    try:
        number = number[1]
    except:
        await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
    else:
        try:
            cur.execute(f'SELECT * FROM Media WHERE number={number}')
            offer = cur.fetchone()
            cur.execute(f'UPDATE Media SET sent=1 WHERE number="{number}"')
        except:
            await call.message.answer('Произошла неизвестная ошибка, связанная с базой данных', reply_markup=menu_kb)
        else:
            await call.message.answer(f'Статус предложения #{number} изменено с {offer[4]} на 1', reply_markup=menu_kb)
            con.commit()


@dp.callback_query_handler(text_startswith='cancel_offer_')
async def offer_me(call: CallbackQuery):
    number = str(call.data).split('cancel_offer_')
    try:
        number = number[1]
    except:
        await call.message.answer('Произошла неизвестная ошибка', reply_markup=menu_kb)
    else:
        try:
            cur.execute(f'SELECT * FROM Media WHERE number={number}')
            offer = cur.fetchone()
            cur.execute(f'UPDATE Media SET sent=-1 WHERE number="{number}"')
        except:
            await call.message.answer('Произошла неизвестная ошибка, связанная с базойд данных',
                                      reply_markup=menu_kb)
        else:
            await call.message.answer(f'Статус предложения #{number} изменено с {offer[4]} на -1',
                                      reply_markup=menu_kb)
            con.commit()


@dp.callback_query_handler(text='offers')
async def offers(call: CallbackQuery):
    cur.execute('SELECT * FROM Media')
    offers = cur.fetchall()
    if not offers:
        await call.message.answer('Предложений нет')
    else:
        temp_kb = InlineKeyboardMarkup()
        for offer in offers:
            number = offer[5]
            temp_kb.add(InlineKeyboardButton(f'Предложение #{number} ({offer[2]})',
                                             callback_data=f'offer_number_{number}'))
        await call.message.answer('Предложения:', reply_markup=temp_kb)


@dp.callback_query_handler(text_startswith='offer_number_')
async def offer_number_(call: CallbackQuery):
    number = call.data.split('offer_number_')
    try:
        number = number[1]
    except:
        await call.message.answer('Произошла неизвестная ошибка')
    else:
        try:
            cur.execute(f'SELECT * FROM Media WHERE number="{number}"')
            offer = cur.fetchone()
            media = offer[1]
            caption = offer[2]
            type = offer[3]
            if type == 'photo':
                try:
                    await bot.send_photo(chat_id=call.from_user.id, photo=media, caption=caption)
                except:
                    await call.message.answer('Произошла неизвестная ошибка')
                else:
                    pass
            elif type == 'video':
                try:
                    await bot.send_video(chat_id=call.from_user.id, video=media, caption=caption)
                except:
                    await call.message.answer('Произошла неизвестная ошибка')
                else:
                    pass
        except:
            await call.message.answer('Произошла неизвестная ошибка')


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
    # executor.start_polling(dp, skip_updates=True, on_startup=schedule_jobs)
