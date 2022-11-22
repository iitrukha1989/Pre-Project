from outlook_templates import plan_graph_db_string, sql_string_pg, sql_string_reestr, sql_string_delay, \
    art_print_start, art_print_end, successfully_send, template_value
import datetime
import pathlib
import sqlite3
import pandas
import pyodbc
import os

# Скрипт по переносу полученных план - графиков (далее по тексту ПГ) и реестров в базы данных
# Разработчик: Ilya Trukhanovich
# Статус: Автоматический запуск, тестирование
# Версия: 1.2 (Введен обновленный шаблон ПГ, учитыающий диапазоны, техническое решение, статус КС-11)
# Расписание запуска: По субботам каждую не четную неделю, после выполнения test_2_xl.py

# ----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции по переносу ПГ (Общее кол-во: 5 функции)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(pathlib.Path.cwd())
tmp_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_templates'
get_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_get'
set_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_set'
back_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_backup'
tmp_date_year = datetime.date.today().strftime("%Y")
week_value = 'W' + str(datetime.datetime.now().isocalendar()[1] - 1)


# 2 этап: Функции
# Функция - 1: Перенос ПГ полученных от ПО в БД
def pg_write_db_1():
    plan_graph_connect = pyodbc.connect(plan_graph_db_string)
    plan_graph_cursor = plan_graph_connect.cursor()
    back_pg_connect = sqlite3.connect(back_dir + r'\database_pg.db')
    back_pg_cursor = back_pg_connect.cursor()
    df_value = pandas.read_excel(
        tmp_dir + r'\Reestr.xlsx', sheet_name='реестр').values.tolist()
    for r_value, d_value, f_value in os.walk(get_dir):
        for file_name in f_value:
            po_name = str(file_name)[23:-14]
            for row_value in df_value:
                if row_value[0] == po_name and row_value[1] == week_value and \
                        (row_value[4] == 'по шаблону' or (row_value[4] == 'с ошибками' and pandas.isna(row_value[8]))):
                    df_value_pg = pandas.read_excel(get_dir + '\\' + file_name,
                                                    sheet_name='ПГ', header=5).values.tolist()
                    for row_pg in df_value_pg:
                        list_pg_option = list()
                        for col_index in range(54):
                            if pandas.isna(row_pg[col_index]):
                                list_pg_option.append('')
                            else:
                                list_pg_option.append(str(row_pg[col_index]))
                        for col_index in range(87, 95):
                            if pandas.isna(row_pg[col_index]):
                                list_pg_option.append('')
                            else:
                                list_pg_option.append(str(row_pg[col_index]))
                        list_pg_option.append(str(week_value) + tmp_date_year)
                        list_pg_option.append(row_value[4])
                        tuple_value_pg = tuple(list_pg_option)
                        select_value = f"insert into [xls].[plan_graph_po] ({sql_string_pg}) values {tuple_value_pg};"
                        plan_graph_cursor.execute(select_value)
                        plan_graph_cursor.commit()
                        select_value = f"insert into database_pg ({sql_string_pg}) values {tuple_value_pg};"
                        back_pg_cursor.execute(select_value)
                        back_pg_connect.commit()
                        list_pg_option.clear()
    plan_graph_connect.close()
    back_pg_connect.close()


# Функция - 2: Перенос ПГ не полученных от ПО в БД
def pg_write_db_2():
    plan_graph_connect = pyodbc.connect(plan_graph_db_string)
    plan_graph_cursor = plan_graph_connect.cursor()
    back_pg_connect = sqlite3.connect(back_dir + r'\database_pg.db')
    back_pg_cursor = back_pg_connect.cursor()
    df_value = pandas.read_excel(
        tmp_dir + r'\Reestr.xlsx', sheet_name='реестр').values.tolist()
    for r_value, d_value, f_value in os.walk(set_dir):
        for file_name in f_value:
            po_name = str(file_name)[23:-14]
            for row_value in df_value:
                if row_value[0] == po_name and row_value[1] == week_value and \
                        row_value[4] in template_value and pandas.isna(row_value[8]):
                    df_value_pg = pandas.read_excel(set_dir + '\\' + file_name,
                                                    sheet_name='ПГ', header=5).values.tolist()
                    for row_pg in df_value_pg:
                        list_pg_option = list()
                        for col_index in range(54):
                            if pandas.isna(row_pg[col_index]):
                                list_pg_option.append('')
                            else:
                                list_pg_option.append(str(row_pg[col_index]))
                        for col_index in range(87, 95):
                            if pandas.isna(row_pg[col_index]):
                                list_pg_option.append('')
                            else:
                                list_pg_option.append(str(row_pg[col_index]))
                        list_pg_option.append(str(week_value) + tmp_date_year)
                        list_pg_option.append(row_value[4])
                        tuple_value_pg = tuple(list_pg_option)
                        select_value = f"insert into [xls].[plan_graph_po] ({sql_string_pg}) values {tuple_value_pg};"
                        plan_graph_cursor.execute(select_value)
                        plan_graph_cursor.commit()
                        select_value = f"insert into database_pg ({sql_string_pg}) values {tuple_value_pg};"
                        back_pg_cursor.execute(select_value)
                        back_pg_connect.commit()
                        list_pg_option.clear()
    plan_graph_connect.close()
    back_pg_connect.close()


# Функция - 3: Перенос реестра в БД
def reestr_wtite_db():
    plan_graph_connect = pyodbc.connect(plan_graph_db_string)
    plan_graph_cursor = plan_graph_connect.cursor()
    back_reestr_connect = sqlite3.connect(back_dir + r'\database_reestr.db')
    back_reestr_cursor = back_reestr_connect.cursor()
    df_value = pandas.read_excel(
        tmp_dir + r'\Reestr.xlsx', sheet_name='реестр').values.tolist()
    list_reestr = list()
    for tmp_row in df_value:
        if tmp_row[1] == week_value:
            for tmp_index in range(10):
                if pandas.isna(tmp_row[tmp_index]):
                    list_reestr.append('')
                else:
                    list_reestr.append(tmp_row[tmp_index])
            tuple_value_reestr = tuple(list_reestr)
            select_value = f"insert into [xls].[reestr_po] ({sql_string_reestr}) values {tuple_value_reestr};"
            plan_graph_cursor.execute(select_value)
            plan_graph_cursor.commit()
            select_value = f"insert into database_reestr ({sql_string_reestr}) values {tuple_value_reestr};"
            back_reestr_cursor.execute(select_value)
            back_reestr_connect.commit()
            list_reestr.clear()
    plan_graph_connect.close()
    back_reestr_connect.close()


# Функция - 4: Перенос реестра задержек в БД
def delay_write_db():
    plan_graph_connect = pyodbc.connect(plan_graph_db_string)
    plan_graph_cursor = plan_graph_connect.cursor()
    back_reestr_connect = sqlite3.connect(
        back_dir + r'\database_reestr_delay.db')
    back_reestr_cursor = back_reestr_connect.cursor()
    df_value = pandas.read_excel(
        tmp_dir + r'\Reestr_delay.xlsx', sheet_name='реестр', header=1).values.tolist()
    list_delay = list()
    for tmp_row in df_value:
        if tmp_row[2] == week_value + tmp_date_year:
            for tmp_index in range(116):
                if pandas.isna(tmp_row[tmp_index]):
                    list_delay.append('')
                else:
                    list_delay.append(tmp_row[tmp_index])
            tuple_value_delay = tuple(list_delay)
            select_value = f"insert into [xls].[reestr_delay_po] ({sql_string_delay}) values {tuple_value_delay}"
            plan_graph_cursor.execute(select_value)
            plan_graph_cursor.commit()
            select_value = f"insert into database_reestr_delay ({sql_string_delay}) values {tuple_value_delay}"
            back_reestr_cursor.execute(select_value)
            back_reestr_connect.commit()
            list_delay.clear()
    plan_graph_connect.close()
    back_reestr_connect.close()


# Функция - 5: Основная функция записи ПГ и реестров в БД
def write_pg():
    pg_write_db_1()
    pg_write_db_2()
    reestr_wtite_db()
    delay_write_db()


# 3 этап: Запуск основной функции/программы
def main_db():
    art_print_start(name_project='Update database')
    data_start = datetime.datetime.now()
    write_pg()
    art_print_end(data_start)
    successfully_send(data_start, name_project='Update database')


if __name__ == '__main__':
    main_db()
