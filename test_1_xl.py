from outlook_templates import iteration_dict, status_db_string, exept_list_po, new_rsr_list, excel_dict, \
    status_select_value, art_print_start, art_print_end, successfully_send, pandas_sql_value_3, function_sort, \
    freezing_select_value_pl, freezing_select_value_cs
import win32com.client
from win32com.client import constants
from win32com.client.gencache import EnsureDispatch
from pathlib import Path
import warnings
import openpyxl
import datetime
import sqlite3
import pandas
import pyodbc
import copy
import os

# Скрипт по формированию шаблонов план - графиков (далее по тексту ПГ) для последующей отправки в адрес ПО
# Разработчик: iitrukha@mtr.ru
# Версия: 1.3 (Введен обновленный шаблон ПГ, учитыающий диапазоны, техническое решение, КС-11, проверка валидности)
# Статус: Автоматический запуск, тестирование
# Расписание запуска: По пятницам каждую четную неделю, в 18:00 по Нск

# ----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции по формированию ПГ (Общее кол-во: 11 функций)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(Path.cwd())
tmp_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_templates'
set_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_set'
back_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_backup'
date_value = datetime.date.today().strftime('%d%m%Y')
date_year_value = datetime.date.today().strftime('%Y')
week_value = 'W' + str(datetime.datetime.now().isocalendar()[1] - 2)
week_year_value = week_value + str(date_year_value)

# 2 этап: Функции
# Функция - 1: Формирование общего статусного отчета по всей РФ из БД ДКРИС
def create_status():
    list_status = list()
    list_status_option = list()
    status_connect = pyodbc.connect(status_db_string)
    status_cursor = status_connect.cursor()
    status_cursor.execute(status_select_value.format(tmp_year_start=int(date_year_value),
                                                     tmp_year_end=int(date_year_value) + 3))
    for row_status in status_cursor:
        if row_status[9] is not None and ((row_status[21] is None and row_status[24] is None) or
                                          (row_status[21] is None and row_status[25] is None) or
                                          (row_status[21] and row_status[24] is None and row_status[25] is None)):
            list_status_option.clear()
            for row_value in row_status:
                if row_value == 'ЦУПРИС Восток':
                    list_status_option.append('Восток')
                    continue
                if row_value == 'ЦУПРИС Запад':
                    list_status_option.append('Запад')
                    continue
                if row_value:
                    list_status_option.append(row_value)
                else:
                    list_status_option.append("")
            list_status.append(list(list_status_option))
    list_status.sort(reverse=True, key=function_sort(2))
    list_status.sort(key=function_sort(1))
    return expend_frost_status(list_status)


# Функция - 2: Формирования списка замороженных объектов по всем РФ из БД ДКРИС
def create_freezing():
    freezing_list = list()
    freezy_connect = pyodbc.connect(status_db_string)
    freezy_cursor = freezy_connect.cursor()
    freezy_cursor.execute(freezing_select_value_pl)
    for row_value in freezy_cursor:
        freezing_list.append(tuple(row_value))
    freezy_cursor.execute(freezing_select_value_cs)
    for row_value in freezy_cursor:
        freezing_list.append(tuple(row_value))
    return freezing_list


# Функция - 3: Исключение из статусного отчета замороженные объектов
def expend_frost_status(list_status):
    result_list_status = copy.deepcopy(list_status)
    freezing_list = create_freezing()
    for row_status in list_status:
        for row_freezing in freezing_list:
            if row_status[1] == row_freezing[0] and row_status[2] == 'Строительство новой площадки (ПК,ПЗ)' and \
                    row_freezing[2] == 'PL' and row_status in result_list_status:
                result_list_status.remove(row_status)
            if row_status[1] == row_freezing[0] and row_status[2] == 'Строительство нового диапазона (ДС)' and \
                    row_freezing[2] != 'PL' and row_status in result_list_status:
                freq_list_1 = sorted(row_status[8].split('/'))
                if '' in freq_list_1:
                    freq_list_1.remove('')
                if 'NI-1800' in freq_list_1:
                    freq_list_1.remove('NI-1800')
                if 'NS-900' in freq_list_1:
                    freq_list_1.remove('NS-900')
                freq_list_2 = list()
                for tmp_row_freezing in freezing_list:
                    if tmp_row_freezing[0] == row_freezing[0] and tmp_row_freezing[2] != 'PL':
                        freq_list_2.append(tmp_row_freezing[2])
                if freq_list_1 == sorted(freq_list_2) or set(freq_list_1).issubset(set(freq_list_2)):
                    result_list_status.remove(row_status)
                freq_list_2.clear()
    return tuple(result_list_status)


# Функция - 4: Формирование списка ПО
def create_po_list(list_status):
    df_value = pandas.read_excel(tmp_dir + r'\name_inn_po.xlsx', sheet_name='name_inn_po').values.tolist()
    po_list = set()
    for tmp_row in list_status:
        inn_value = tmp_row[7][-11:-1]
        for row_value in df_value:
            if inn_value == str(row_value[1]) and tmp_row[7] not in exept_list_po:
                po_list.add(row_value[0])
                break
    return sorted(po_list)


# Функция - 5: Формирование ПГ
def create_pg():
    list_status = create_status()
    po_list = create_po_list(list_status)
    for po_value in po_list:
        create_pg_option(po_value, list_status)


# Функция - 6: Опционная функция формирование ПГ из статусного отчета по определенному ПО
def create_pg_option(po_value, list_status):
    warnings.simplefilter(action='ignore', category=UserWarning)
    check_status = 0
    check_pl_list = list()
    book_value = openpyxl.load_workbook(tmp_dir + r'\Templates.xlsx')
    sheet_value = book_value.active
    tmp_index = 7
    for tmp_row in list_status:
        inn_value = po_value[-11:-1]
        inn_value_status = tmp_row[7][-11:-1]
        if inn_value == inn_value_status and tmp_row[2] in new_rsr_list:
            for key_value in iteration_dict:
                if key_value <= 7:
                    sheet_value[tmp_index][key_value].value = tmp_row[iteration_dict[key_value]]
                elif 7 < key_value < 48:
                    sheet_value[tmp_index][key_value].value = tmp_row[iteration_dict[key_value]][:10]
                elif key_value == 48:
                    if iteration_dict[key_value][1] is None:
                        sheet_value[tmp_index][key_value].value = tmp_row[iteration_dict[key_value][0]][:10]
                    else:
                        sheet_value[tmp_index][key_value].value = tmp_row[iteration_dict[key_value][1]][:10]
            sheet_value[tmp_index][6].value = po_value
            sheet_value[tmp_index][53].value = tmp_row[iteration_dict[53]]
            check_database(sheet_value, po_value, tmp_index, number_pl=tmp_row[1])
            check_not_req(sheet_value, tmp_index)
            tmp_index = check_duplicate(sheet_value, tmp_index, check_pl_list)
            if sheet_value[tmp_index][7].value in ('NI-1800/', 'NS-900/'):
                sheet_value.delete_rows(tmp_index, 1)
            else:
                check_pl_list.append(tmp_row[1])
                check_status = 1
                tmp_index += 1
    sheet_value.delete_rows(tmp_index, 1000)
    book_value.save(set_dir + str(r'\План-график БС ПАО МТС_' + po_value + '_' + date_value + '.xlsx'))
    book_value.close()
    update_formula(po_value, tmp_index)
    if check_status == 0:
        os.remove(set_dir + str(r'\План-график БС ПАО МТС_' + po_value + '_' + date_value + '.xlsx'))


# Функция - 7: Сверка/дополнение сформированного ПГ с данными из локальной БД
def check_database(sheet_value, po_value, tmp_index, number_pl):
    warnings.simplefilter(action='ignore', category=UserWarning)
    db_connect = sqlite3.connect(back_dir + r'\database_pg.db')
    df_value = pandas.read_sql_query(pandas_sql_value_3.format(pl=number_pl, po=po_value, week_year=week_year_value),
                                     db_connect).values.tolist()
    for row_value in df_value:
        if sheet_value[tmp_index][1].value == row_value[1]:
            for col_index in range(9, 50, 2):
                if pandas.isna(row_value[col_index]):
                    sheet_value[tmp_index][col_index].value = ''
                else:
                    sheet_value[tmp_index][col_index].value = check_database_option(db_value=row_value[col_index])
            for col_index in range(10, 51, 2):
                if sheet_value[tmp_index][col_index].value is None or sheet_value[tmp_index][col_index].value == '':
                    if pandas.isna(row_value[col_index]):
                        sheet_value[tmp_index][col_index].value = ''
                    else:
                        sheet_value[tmp_index][col_index].value = check_database_option(db_value=row_value[col_index])
            if pandas.isna(row_value[51]):
                sheet_value[tmp_index][51].value = ''
            else:
                sheet_value[tmp_index][51].value = row_value[51]
            if pandas.isna(row_value[52]) or row_value[52] == 'не отражен в план-графике':
                sheet_value[tmp_index][52].value = ''
            else:
                sheet_value[tmp_index][52].value = row_value[52]
    db_connect.close()


# Функция - 8: Опционная функция проверки валидности данных записанных в локульную БД
def check_database_option(db_value):
    res_value = ''
    if db_value in ['не треб', 'не треб.', 'не требуется']:
        res_value = 'не треб'
    else:
        try:
            datetime.datetime.strptime(db_value, '%Y-%m-%d %H:%M:%S')
            res_value = db_value[8:10] + '.' + db_value[5:7] + '.' + db_value[:4]
        except:
            pass
        try:
            datetime.datetime.strptime(db_value, '%d.%m.%Y %H:%M:%S')
            res_value = db_value[:2] + '.' + db_value[3:5] + '.' + db_value[6:10]
        except:
            pass
        try:
            datetime.datetime.strptime(db_value, '%d.%m.%Y')
            res_value = db_value[:2] + '.' + db_value[3:5] + '.' + db_value[6:10]
        except:
            pass
    return res_value


# Функция - 9: Опционная функция автоматического заполнения некоторых этапов не требующих выполнения
def check_not_req(sheet_value, tmp_index):
    if sheet_value[tmp_index][2].value == 'CS':
        for cs_index in [13, 17, 27]:
            if (sheet_value[tmp_index][cs_index].value == '' and sheet_value[tmp_index][cs_index + 1].value == '') or \
                    (sheet_value[tmp_index][cs_index].value is None and
                     sheet_value[tmp_index][cs_index + 1].value is None):
                sheet_value[tmp_index][cs_index].value = 'не треб'
    if sheet_value[tmp_index][2].value == 'RT':
        for cs_index in [13, 27]:
            if (sheet_value[tmp_index][cs_index].value == '' and sheet_value[tmp_index][cs_index + 1].value == '') or \
                    (sheet_value[tmp_index][cs_index].value is None and
                     sheet_value[tmp_index][cs_index + 1].value is None):
                sheet_value[tmp_index][cs_index].value = 'не треб'
    if sheet_value[tmp_index][2].value == 'GF':
        for cs_index in [15]:
            if (sheet_value[tmp_index][cs_index].value == '' and sheet_value[tmp_index][cs_index + 1].value == '') or \
                    (sheet_value[tmp_index][cs_index].value is None and
                     sheet_value[tmp_index][cs_index + 1].value is None):
                sheet_value[tmp_index][cs_index].value = 'не треб'


# Функция - 10: Опционная функция объединение дублирующих строк в ПГ
def check_duplicate(sheet_value, tmp_index, check_pl_list):
    if sheet_value[tmp_index][1].value in check_pl_list:
        check_flag = 0
        set_1 = set(sheet_value[tmp_index][7].value.split('/'))
        for row_value in range(7, tmp_index):
            if sheet_value[row_value][1].value == sheet_value[tmp_index][1].value and \
                    sheet_value[row_value][2].value == sheet_value[tmp_index][2].value:
                set_2 = set(sheet_value[row_value][7].value.split('/'))
                if set_1.issubset(set_2) or set_2.issubset(set_1) or \
                        (sheet_value[row_value][2].value == 'GF' and sheet_value[tmp_index][2].value == 'GF') or \
                        (sheet_value[row_value][2].value == 'RT' and sheet_value[tmp_index][2].value == 'RT'):
                    check_flag = 1
                    set_1.update(set_2)
                    sheet_value[row_value][7].value = '/'.join(sorted(list(set_1)))
                    for col_index in range(9, 51):
                        if sheet_value[row_value][col_index].value is None or \
                                sheet_value[row_value][col_index].value == '':
                            sheet_value[row_value][col_index].value = sheet_value[tmp_index][col_index].value
                    break
        if check_flag == 1:
            sheet_value.delete_rows(tmp_index, 1)
            tmp_index -= 1
    return tmp_index


# Функция - 11: Функция по обновлению формул корректности план-графиков
def update_formula(po_value, tmp_index):
    excel_api = win32com.client.Dispatch("Excel.Application")
    excel_api.Visible = False
    book_value = excel_api.Workbooks.Open(set_dir +
                                          str(r'\План-график БС ПАО МТС_' + po_value + '_' + date_value + '.xlsx'))
    sheet_value = book_value.ActiveSheet
    for key_value in excel_dict:
        sheet_value.Range(key_value).Formula = excel_dict[key_value]
    if tmp_index > 8:
        xl = EnsureDispatch('Excel.Application')
        xl.Range("BC7:BY7").Select()
        xl.Selection.AutoFill(Destination=sheet_value.Range(f"BC7:BY{tmp_index - 1}"), Type=constants.xlFillDefault)
    book_value.Close(True)


# 3 этап: Запуск основной функции/программы
def main_xl_1():
    art_print_start(name_project='Create PG')
    data_start = datetime.datetime.now()
    create_pg()
    art_print_end(data_start)
    successfully_send(data_start, name_project='Create PG')


if __name__ == '__main__':
    main_xl_1()
