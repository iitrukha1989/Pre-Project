from outlook_templates import sql_table_backup, sql_table_header, art_print_start, art_print_end, successfully_send, \
    plan_graph_db_string
from pathlib import Path
import datetime
import pyodbc
import csv


# Скрипт по созданию backup файлов: ПГ, реестр, реестр_задержек, в формате .csv
# Разработчик: Ilya Trukhanovich
# Статус: Разработка, тестирования
# Требуется: Дописать функционал архивирования полученных backup файлов с последующей отправкой на почту разработчику

# ----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции по формированию ПГ (Общее кол-во: 1 функция)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(Path.cwd())
back_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_backup'


# 2 этап: Функции
# Функция - 1: Выгрузка данных из БД SQL (ПГ, реестр, реестр_задержек) и сохранение в файл формата .csv
def csv_create():
    for key_value in sql_table_backup:
        dkris_connect = pyodbc.connect(plan_graph_db_string)
        dkris_cursor = dkris_connect.cursor()
        select_value = f"select * from {key_value}"
        dkris_cursor.execute(select_value)
        with open(back_dir + r'\\' + sql_table_backup[key_value] + '.csv', 'w', newline='') as csv_file:
            writer = csv.writer(csv_file, delimiter=";")
            writer.writerow(sql_table_header[key_value])
            for tmp_row in dkris_cursor:
                writer.writerow(tmp_row)


# 3 этап: Запуск основной функции/программы
def main_bg():
    art_print_start(name_project='Update backup')
    data_start = datetime.datetime.now()
    csv_create()
    art_print_end(data_start)
    successfully_send(data_start, name_project='Update backup')


if __name__ == '__main__':
    main_bg()
