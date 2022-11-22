from outlook_templates import art_print_start, art_print_end, successfully_send, html_value_3_1, html_value_3_2, \
    html_value_3_3, html_value_3_4, html_value_3_5, html_value_3_6, html_value_3_7, html_value_3_8, html_value_3_9, \
    html_value_3_10
import win32com.client as client
import pythoncom
import datetime
import openpyxl
import pathlib
import pandas
import os

# Скрипт по рассылке запросов в адрес ПО с вложенными план-графиками (далее по тексту ПГ)
# Разработчик: iitrukha@mtr.ru
# Версия: 1.3 (Используется адаптивное формирование текста письма, в качестве вложения используется новый шаблон ПГ)
# Статус: Автоматический запуск, тестирование
# Расписание запуска: По пятницам каждую четную неделю, после выполнения test_3_xl.py

# ----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции по отправке ПГ в адрес ПО (Общее кол-во: 5 функций)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(pathlib.Path.cwd())
tmp_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_templates'
set_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_set'
date_value = datetime.datetime.now().strftime('%d%m%Y')
week_value = 'W' + str(datetime.datetime.now().isocalendar()[1])


# 2 этап:
# Функция - 1: Получение максимального значения строк реестров
def rows_reestrs():
    return len(pandas.read_excel(tmp_dir + r'\Reestr.xlsx', sheet_name='реестр').values.tolist()) + 1


# Функция - 2: Отправка писем/ПГ в адрес ПО
def send_mail():
    df_value_contact = pandas.read_excel(tmp_dir + r'\name_inn_po.xlsx', sheet_name='name_inn_po').values.tolist()
    for tmp_root_set, tmp_dir_set, tmp_file_set in os.walk(set_dir):
        for book_name in tmp_file_set:
            if book_name.endswith('.zip'):
                inn_po = pathlib.Path(book_name).stem[-20:-10]
                name_po = str(book_name)[23:-13]
                df_value_pg = pandas.read_excel(set_dir + f'\План-график БС ПАО МТС_{name_po}_{date_value}.xlsx',
                                                sheet_name='ПГ', header=5).values.tolist()
                po_dict = {87: list(), 90: list(), 91: list(), 92: list(), 93: list(), 94: list()}
                for row_value in df_value_pg:
                    for tmp_index in [87, 90, 91, 92, 93, 94]:
                        if pandas.isna(row_value[tmp_index]) is False:
                            po_dict[tmp_index].append(row_value[1])
                for tmp_row in df_value_contact:
                    if tmp_row[1] == int(inn_po):
                        name_po = tmp_row[0]
                        email_po = tmp_row[2]
                        send_mail_option(book_name, name_po, email_po, po_dict)


# Функция - 3: Опционная функция отправки писем/ПГ в адрес ПО
def send_mail_option(book_name, name_po, email_po, po_dict):
    outlook_value = client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    send_account = None
    for account in outlook_value.Session.Accounts:
        if account.DisplayName == "cupris-vostok@mts.ru":
            send_account = account
            break
    message_value = outlook_value.CreateItem(0)
    message_value._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    message_value.Subject = f"План-график БС ПАО МТС_{name_po}"
    create_html(message_value, po_dict)
    message_value.Attachments.Add(set_dir + '\\' + book_name)
    message_value.To = email_po
    message_value.CC = 'data.cupris_west@mts.ru; cupris-vostok@mts.ru; cupris_west@mts.ru'
    message_value.Send()
    log_reestr(name_po)
    os.remove(set_dir + '\\' + book_name)


# Функция - 4: Формирования адаптивного текста письма для рассылки в адрес ПО
def create_html(message_value, po_dict):
    check_sum = 0
    message_value.HTMLBody = html_value_3_1.format(date_value_1=(datetime.datetime.now() +
                                                                 datetime.timedelta(days=5)).strftime("%d.%m.%Y"))
    for tmp_index in [87, 90, 91, 92, 93, 94]:
        check_sum += int(len(po_dict[tmp_index]))
    if check_sum != 0:
        message_value.HTMLBody += html_value_3_2
    if len(po_dict[87]) != 0:
        message_value.HTMLBody += html_value_3_3.format(count_to=len(po_dict[87]), list_to=', '.join(po_dict[87]))
    if len(po_dict[90]) != 0:
        message_value.HTMLBody += html_value_3_4.format(count_smr=len(po_dict[90]), list_smr=', '.join(po_dict[90]))
    if len(po_dict[91]) != 0:
        message_value.HTMLBody += html_value_3_5.format(count_smrg=len(po_dict[91]), list_smrg=', '.join(po_dict[91]))
    if len(po_dict[92]) != 0:
        message_value.HTMLBody += html_value_3_6.format(count_check=len(po_dict[92]), list_check=', '.join(po_dict[92]))
    if len(po_dict[93]) != 0:
        message_value.HTMLBody += html_value_3_7.format(count_ks=len(po_dict[93]), list_ks=', '.join(po_dict[93]))
    if len(po_dict[94]) != 0:
        message_value.HTMLBody += html_value_3_8.format(count_aop=len(po_dict[94]), list_aop=', '.join(po_dict[94]))
    if check_sum != 0:
        message_value.HTMLBody += html_value_3_9
    message_value.HTMLBody += html_value_3_10.format(date_value_2=(datetime.datetime.now() +
                                                                   datetime.timedelta(days=6)).strftime("%d.%m.%Y"))


# Функция - 5: Заполнение реестра по отправкам писем в адрес ПО
def log_reestr(name_po):
    row_index = rows_reestrs()
    book_value = openpyxl.load_workbook(tmp_dir + r'\Reestr.xlsx')
    sheet_value = book_value.active
    sheet_value[row_index + 1][0].value = name_po
    sheet_value[row_index + 1][1].value = week_value
    sheet_value[row_index + 1][2].value = datetime.date.today().strftime("%d.%m.%Y")
    book_value.save(tmp_dir + r'\Reestr.xlsx')
    book_value.close()


# 3 этап: Запуск основной функции/программы
def main_send_1():
    art_print_start(name_project='Send email')
    data_start = datetime.datetime.now()
    send_mail()
    art_print_end(data_start)
    successfully_send(data_start, name_project='Send email')


if __name__ == '__main__':
    main_send_1()
