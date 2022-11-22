from outlook_templates import sql_tuple_1, sql_tuple_2, status_db_string, art_print_start, art_print_end, \
    successfully_send, function_sort, dict_stage_pg, status_select_value_po, iteration_dict_upd, new_rsr_list, \
    html_value_4_1, html_value_4_2, html_value_4_3, html_value_4_4, html_value_4_5, html_value_4_6
from win32com import client
import pythoncom
import openpyxl
import tabulate
import datetime
import pathlib
import sqlite3
import pyodbc
import pandas
import os

# Скрипт по отправке сводных данных по ПГ в адрес ОРС
# Разработчик: iitrukha@mts.ru
# Версия: 1.3 (Используется адаптивное формирование текста письма для рассылки в адрес ОРС)
# Статус: Автоматический запуск, тестирование
# Расписание запуска: По понедельника каждую четную неделю, в 09:00 по Нск

# -----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции автоматической рассылке (Общее кол-во: 12 функций)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(pathlib.Path.cwd())
tmp_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_templates'
set_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_set'
get_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_get'
back_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_backup'
date_year_value = datetime.datetime.today().strftime("%Y")
week_value = 'W' + str(datetime.datetime.now().isocalendar()[1] - 2)
week_year_value = week_value + str(date_year_value)


# 2 этап: Функции
# Функция - 1: Обновление ПГ в части учета фактических дат по ключевым этапам на текущую дату
def update_pg():
    for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
        for file_name in get_f_value:
            po_name = str(file_name)[23:-14]
            for set_r_value, set_d_value, set_f_value in os.walk(set_dir):
                for file_name_1 in set_f_value:
                    po_name_1 = str(file_name_1)[23:-14]
                    if po_name_1 == po_name:
                        status_list = create_status(po_name)
                        book_value = openpyxl.open(set_dir + '\\' + file_name_1)
                        sheet_value = book_value.active
                        row_index = 7
                        while sheet_value[row_index][1].value:
                            for row_value in status_list:
                                if sheet_value[row_index][1].value == row_value[1] and \
                                        sheet_value[row_index][2].value == row_value[3] and \
                                        sheet_value[row_index][7].value == row_value[8] and \
                                        row_value[2] in new_rsr_list:
                                    for col_index in iteration_dict_upd.keys():
                                        sheet_value[row_index][col_index].value = row_value[iteration_dict_upd[col_index]]
                            row_index += 1
                        book_value.save(set_dir + '\\' + file_name_1)
                        book_value.close()


# Функция - 2: Формирование выгрузки из БД ДКРИС по определенному ПО, за текущую дату
def create_status(po_name):
    list_status = list()
    list_status_option = list()
    status_connect = pyodbc.connect(status_db_string)
    status_cursor = status_connect.cursor()
    status_cursor.execute(status_select_value_po.format(tmp_year_start=int(date_year_value),
                                                        tmp_year_end=int(date_year_value) + 3, po_name=po_name))
    for tmp_status in status_cursor:
        list_status_option.clear()
        for col_index in range(27):
            if col_index < 24:
                if tmp_status[col_index]:
                    list_status_option.append(tmp_status[col_index])
                else:
                    list_status_option.append("")
            if col_index == 24:
                if tmp_status[24] is not None:
                    list_status_option.append(tmp_status[24])
                else:
                    list_status_option.append(tmp_status[25])
            if col_index == 26:
                if tmp_status[26] == 'ЦУПРИС Восток':
                    list_status_option.append('Восток')
                elif tmp_status[26] == 'ЦУПРИС Запад':
                    list_status_option.append('Запад')
                else:
                    list_status_option.append(tmp_status[26])
        list_status.append(list(list_status_option))
    list_status.sort(reverse=True, key=function_sort(2))
    list_status.sort(key=function_sort(1))
    return tuple(list_status)


# Функция - 3: Формирование списка регионов
def create_region_list():
    region_list = set()
    db_connect = sqlite3.connect(back_dir + r'\database_pg.db')
    select_value = """select region from database_pg where number_week_year = '{week_year_value}'
                   and (status_week = 'по шаблону' or status_week = 'с ошибками')"""
    sql_value = pandas.read_sql_query(select_value.format(week_year_value=week_year_value), db_connect)
    df_value = pandas.DataFrame(sql_value, columns=['region']).values.tolist()
    for tmp_value in df_value:
        region_list.add(*tmp_value)
    return sorted(region_list)


# Функция - 4: Формирование списка расхождений по фактическим датам в БД и ПГ
def create_info_diff():
    dict_diff = dict()
    for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
        for get_file in get_f_value:
            name_po_get = str(get_file)[23:-14]
            get_df_value = pandas.read_excel(get_dir + '\\' + get_file, sheet_name='ПГ', header=5).values.tolist()
            for set_r_value, set_d_value, set_f_value in os.walk(set_dir):
                for set_file in set_f_value:
                    name_po_set = str(set_file)[23:-14]
                    if name_po_get == name_po_set:
                        set_df_value = pandas.read_excel(set_dir + '\\' + set_file, sheet_name='ПГ',
                                                         header=5).values.tolist()
                        for set_row in set_df_value:
                            for get_row in get_df_value:
                                if set_row[1] == get_row[1]:
                                    for tmp_index in dict_stage_pg.keys():
                                        if pandas.isna(set_row[tmp_index]) and pandas.isna(get_row[tmp_index]) is False:
                                            key_dict = (set_row[0], set_row[1])
                                            date_value = get_row[tmp_index]
                                            if isinstance(get_row[tmp_index], datetime.datetime):
                                                date_value = str(get_row[tmp_index].strftime('%d.%m.%Y'))
                                            dict_diff[key_dict] = (tmp_index, date_value, get_row[6])
    return dict_diff


# Функция - 5: Формирование списка ПО по которым получен ПГ за отчетную неделю (по шаблону или с ошибками)
def create_info_po():
    po_get = list()
    df_value = pandas.read_excel(tmp_dir + r'\Reestr.xlsx', sheet_name='реестр').values.tolist()
    for tmp_value in df_value:
        if tmp_value[1] == week_value and tmp_value[4] == 'по шаблону' or (tmp_value[4] == 'с ошибками' and
                                                                           pandas.isna(tmp_value[8])):
            po_get.append(tmp_value[0])
    return tuple(po_get)


# Функция - 7: Функция рассылки писем в адрес ОРС
def send_mail():
    update_pg()
    region_list = create_region_list()
    dict_diff = create_info_diff()
    po_get = create_info_po()
    for region_name in region_list:
        send_mail_option(region_name, dict_diff, po_get)


# Функция - 8: Опционная функция формирования вложения, текста письма, и отправки в адрес ОРС
def send_mail_option(region_name, dict_diff, po_get):
    outlook_value = client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    for account in outlook_value.Session.Accounts:
        if account.DisplayName == "cupris-vostok@mts.ru":
            send_account = account
            break
    message_value = outlook_value.CreateItem(0)
    message_value._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    message_value.Subject = f"План-график по объектам сети радиодоступа ({region_name}), итоги за {week_value}"
    create_html(message_value, region_name, dict_diff, po_get)
    message_value.Attachments.Add(set_dir + '\\' + f'План-график ({region_name}).xlsx')
    message_value.To = create_region_contact(region_name)
    message_value.CC = 'cupris_west@mts.ru; cupris-vostok@mts.ru'
    message_value.Send()
    os.remove(set_dir + '\\' + f'План-график ({region_name}).xlsx')


# Функция - 9: Формирования адаптивного текста письма для рассылки в адрес ОРС
def create_html(message_value, region_name, dict_diff, po_get):
    po_list_get = set()
    po_list_1 = list()
    po_list_2 = list()
    po_list_3 = list()
    po_list_4 = list()
    po_list_1.append(('Номер площадки', 'Этап строительства', 'Дата из ПГ', 'Наименование ПО'))
    po_list_2.append(('Номер площадки', 'Адрес', 'Наименование ПО'))
    po_list_3.append(('Номер площадки', 'Адрес', 'Наименование ПО'))
    po_list_4.append(('Номер площадки', 'Наименование ПО', 'Блок-Фактор'))
    book_value = openpyxl.open(back_dir + r'\Templates_region.xlsx')
    sheet_value = book_value.active
    row_index = 7
    for tmp_value in create_dataframe(region_name):
        if tmp_value[52] == 'не отражен в план-графике':
            po_list_2.append((tmp_value[1], tmp_value[4], tmp_value[6]))
        if tmp_value[52] == 'не передавалось в работу':
            po_list_3.append((tmp_value[1], tmp_value[4], tmp_value[6]))
        if tmp_value[52] not in ('не передавалось в работу', 'не отражен в план-графике') and tmp_value[52] != '':
            po_list_4.append((tmp_value[1], tmp_value[6], tmp_value[52]))
        if tmp_value[6] in po_get:
            po_list_get.add(tmp_value[6])
        for col_index in range(54):
            sheet_value[row_index][col_index].value = tmp_value[col_index]
        row_index += 1
    for tmp_key, tmp_value in sorted(dict_diff.items(), key=function_sort(1)):
        if region_name in tmp_key:
            po_list_1.append((tmp_key[1], dict_stage_pg[tmp_value[0]], tmp_value[1], tmp_value[2]))
    message_value.HTMLBody = html_value_4_1.format(list_po=', '.join(po_list_get), week_value=week_value[1:])
    if len(po_list_4) > 1:
        message_value.HTMLBody += html_value_4_2.format(table_4=tabulate.tabulate(po_list_4, headers='firstrow',
                                                                                  tablefmt='html', stralign='left'))
    if len(po_list_1) > 1:
        message_value.HTMLBody += html_value_4_3.format(table_1=tabulate.tabulate(po_list_1, headers='firstrow',
                                                                                  tablefmt='html', stralign='left'))
    if len(po_list_2) > 1:
        message_value.HTMLBody += html_value_4_4.format(table_2=tabulate.tabulate(po_list_2, headers='firstrow',
                                                                                  tablefmt='html', stralign='left'))
    if len(po_list_3) > 1:
        message_value.HTMLBody += html_value_4_5.format(table_3=tabulate.tabulate(po_list_3, headers='firstrow',
                                                                                  tablefmt='html', stralign='left'))
    message_value.HTMLBody += html_value_4_6
    book_value.save(set_dir + '\\' + f'План-график ({region_name}).xlsx')
    book_value.close()


# Функция - 10: Создание DataFrame из локальной БД по определнному региону
def create_dataframe(region_name):
    db_connect = sqlite3.connect(back_dir + r'\database_pg.db')
    select_value = "select * from database_pg where number_week_year = '{week_year_value}' and " \
                   "region = '{region}' and (status_week = 'по шаблону' or status_week = 'с ошибками')"
    sql_value = pandas.read_sql_query(select_value.format(week_year_value=week_year_value,
                                                          region=region_name), db_connect)
    df_value = pandas.DataFrame(sql_value, columns=(sql_tuple_1 + sql_tuple_2)).values.tolist()
    return df_value


# Функция - 11: Формирование списка контактных данных ОРС
def create_region_contact(region_name):
    df_value = pandas.read_excel(tmp_dir + r'\Contact_ors.xlsx', sheet_name='ТД_ОРС').values.tolist()
    for tmp_value in df_value:
        if tmp_value[9] == region_name:
            return str(tmp_value[14] + '; ' + tmp_value[21])


# Функция - 12: Очистка папки outlook_set/get от ПГ
def clear_dir():
    for root_value, dirs_value, files_value in os.walk(set_dir):
        for file_name in files_value:
            os.remove(set_dir + '\\' + file_name)
    for root_value, dirs_value, files_value in os.walk(get_dir):
        for file_name in files_value:
            os.remove(get_dir + '\\' + file_name)


# 3 этап: Запуск основной функции/программы
def main_send_3():
    art_print_start(name_project='Send mail')
    date_start = datetime.datetime.now()
    send_mail()
    clear_dir()
    art_print_end(date_start)
    successfully_send(date_start, name_project='Send mail')


if __name__ == '__main__':
    main_send_3()
