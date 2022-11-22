from outlook_templates import art_print_start, art_print_end, successfully_send
import win32com.client as client
from pathlib import Path
import pythoncom
import warnings
import openpyxl
import datetime
import pandas
import os
# import win32timezone


# Скрипт получения от ПО план-графиков (далее по тексту ПГ), из почтового сервера
# Разработчик: Ilya Trukhanovich
# Версия: 1.2
# Статус: Автоматический запуск, тестирование
# Расписание запуска: По четвергам и субботам каждую не четную неделю, в 09:00 по Нск.

# ----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции по получению ПГ (Общее кол-во: 4 функции)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(Path.cwd())
get_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_get'
set_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_set'
tmp_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_templates'
tmp_date_day_1 = datetime.date.today().strftime("%d")
week_value = 'W' + str(datetime.datetime.now().isocalendar()[1] - 1)


# 2 этап:
# Функция - 1: Получение писем/ПГ от ПО
def get_mail():
    input_index = 1
    warnings.simplefilter(action='ignore', category=UserWarning)
    incorrect_list = list()
    check_error = 0
    current_day = datetime.datetime.isoweekday(datetime.datetime.now())
    if (current_day == 4):
        limit_day = (datetime.datetime.today() - datetime.timedelta(days=7))
    else:
        limit_day = (datetime.datetime.today() - datetime.timedelta(days=9))
    outlook = client.Dispatch("Outlook.Application", pythoncom.CoInitialize()).GetNamespace("MAPI")
    inbox = outlook.Folders("cupris-vostok@mts.ru").Folders("Входящие").Folders("План-графики").Items
    messages = inbox.GetLast()
    while messages:
        if messages.Class == 43:
            sender_day = str(messages.SentOn)[8:10] + '.' + str(messages.SentOn)[5:7] + '.' + str(messages.SentOn)[:4]
            # print(messages.SenderEmailAddress, sender_day, limit_day.strftime("%d.%m.%Y"))
            if datetime.datetime.strptime(sender_day, "%d.%m.%Y") < limit_day:
                check_error += 1
                if check_error == 10:
                    break
                messages = inbox.GetPrevious()
                continue
            tmp_attachments = messages.Attachments
            for tmp_attachment in tmp_attachments:
                tmp_filename, tmp_file_extend = os.path.splitext(str(tmp_attachment))
                if tmp_file_extend in ['.rar', '.zip', '.xls', 'xlsb', 'xlsm']:
                    log_reestr(sender_day, sender_value=messages.SenderEmailAddress)
                if tmp_file_extend == '.xlsx':
                    tmp_attachment.SaveAsFile(get_dir + '\\' + str(input_index) + '_' + sender_day[:2] +
                                              sender_day[3:5] + '_' + str(tmp_attachment))
                    try:
                        openpyxl.open(get_dir + '\\' + str(input_index) + '_' + sender_day[:2] +
                                              sender_day[3:5] + '_' + str(tmp_attachment))
                    except:
                        incorrect_list.append((str(input_index) + '_' + sender_day[:2] +
                                              sender_day[3:5] + '_' + str(tmp_attachment)))
                        log_reestr(sender_day,  sender_value=messages.SenderEmailAddress)
                    input_index += 1
        messages = inbox.GetPrevious()
    clear_incorrect(incorrect_list)


# Функция - 2: Опционная функция проверки полученных ПГ с регистрацией в реестре
def log_reestr(sender_day, sender_value):
    df_value = pandas.read_excel(tmp_dir + r'\name_inn_po.xlsx', sheet_name='name_inn_po').values.tolist()
    for tmp_value in df_value:
        sender_list = list()
        if pandas.isna(tmp_value[2]) is False:
            sender_list = tmp_value[2].split(';')
        if sender_value in sender_list:
            book_value = openpyxl.open(tmp_dir + r'\Reestr.xlsx')
            sheet_value = book_value.active
            index_row = 1
            while sheet_value[index_row][0].value:
                index_row += 1
            for tmp_index in range(1, index_row):
                if sheet_value[tmp_index][0].value == tmp_value[0] and \
                        sheet_value[tmp_index][1].value == week_value and \
                        sheet_value[tmp_index][4].value is None:
                    sheet_value[tmp_index][3].value = sender_day
                    sheet_value[tmp_index][4].value = 'не по шаблону'
                    if datetime.datetime.isoweekday(datetime.datetime.now()) == 4:
                        sheet_value[tmp_index][5].value = datetime.date.today().strftime("%d.%m.%Y")
                        sheet_value[tmp_index][8].value = 'да'
                        sheet_value.insert_rows(tmp_index + 1, 1)
                        sheet_value[tmp_index + 1][0].value = sheet_value[tmp_index][0].value
                        sheet_value[tmp_index + 1][1].value = sheet_value[tmp_index][1].value
                        sheet_value[tmp_index + 1][2].value = datetime.date.today().strftime("%d.%m.%Y")
            book_value.save(tmp_dir + r'\Reestr.xlsx')
            book_value.close()


# Функция - 3: Итоговая проверка полученных план - графиков от ПО на предмет соответствия шаблону
def check_pg():
    for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
        for file_name in get_f_value:
            check_flag = 0
            try:
                openpyxl.open(get_dir + '\\' + file_name).active[7][6].value
            except:
                check_flag = 1
                os.remove(get_dir + '\\' + file_name)
            if check_flag == 0:
                po_name = openpyxl.open(get_dir + '\\' + file_name).active[7][6].value
                cypris_value = openpyxl.open(get_dir + '\\' + file_name).active[4][53].value
                check_po = 1
                for set_r_value, set_d_value, set_f_value in os.walk(set_dir):
                    for file_set in set_f_value:
                        name_po_set = str(file_set)[23:-14]
                        if name_po_set == po_name:
                            check_po = 0
                if cypris_value != 'ЦУПРИС':
                    check_po = 1
                if check_po == 1:
                    os.remove(get_dir + '\\' + file_name)


# Функция - 4: Первичная очистка не корректных ПГ
def clear_incorrect(incorrect_list):
    for tmp_value in incorrect_list:
        try:
            os.remove(get_dir + '\\' + tmp_value)
        except:
            pass


# 3 этап: Запуск основной функции/программы
def main_get_1():
    art_print_start(name_project='Get email')
    data_start = datetime.datetime.now()
    get_mail()
    check_pg()
    art_print_end(data_start)
    successfully_send(data_start, name_project='Get email')


if __name__ == '__main__':
    main_get_1()
