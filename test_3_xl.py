from outlook_templates import check_dict, art_print_start, art_print_end, successfully_send, error_value_list, \
    error_value_date
from openpyxl.worksheet.datavalidation import DataValidation
from zipencrypt import ZipFile
import openpyxl
import pathlib
import datetime
import pandas
import os

# Скрипт по проверке полученных план - графиков (далее по тексту ПГ) на предмет отставания по основным этапам
# Разработчик: Ilya Trukhanovich
# Версия: 1.2 (Введен обновленный шаблон ПГ, учитыающий диапазоны, техническое решение, просрочки по статусу КС-11)
# Статус: Автоматический запуск, тестирование
# Расписание запуска: По пятницам каждую четную неделю, после выполнения test_1_xl.py

# ----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции по проверки ПГ (Общее кол-во: 11 функций)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(pathlib.Path.cwd())
tmp_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_templates'
get_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_get'
set_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_set'
back_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_backup'
date_year_value = datetime.date.today().strftime("%Y")
week_value = 'W' + str(datetime.datetime.now().isocalendar()[1])


# 2 этап: Функции
# Функция - 1: Получение максимального значения строк реестров
def rows_reestrs():
    df_value = pandas.read_excel(
        tmp_dir + r'\Reestr_delay.xlsx', sheet_name='реестр').values.tolist()
    return len(df_value) + 2


# Функция - 2: Формирование списка регионов в которых работает определенный ПО
def create_region_list(file_name):
    region_list = set()
    df_value = pandas.read_excel(
        set_dir + '\\' + file_name, sheet_name='ПГ', header=5).values.tolist()
    for row_value in df_value:
        region_list.add(row_value[0])
    return tuple(region_list)


# Функция - 3: Проверка валидности фактических дат
def check_date_value(date_value_1, date_value_2):
    if isinstance(date_value_1, datetime.datetime):
        date_value_1 = str(date_value_1.strftime('%d.%m.%Y'))
    if isinstance(date_value_1, str) and date_value_1 is not None:
        date_value_1 = date_value_1[:2] + '.' + \
            date_value_1[3:5] + '.' + date_value_1[6:10]
    if isinstance(date_value_2, datetime.datetime):
        date_value_2 = str(date_value_2.strftime('%d.%m.%Y'))
    if isinstance(date_value_2, str) and date_value_2 is not None:
        date_value_2 = date_value_2[:2] + '.' + \
            date_value_2[3:5] + '.' + date_value_2[6:10]
    return date_value_1, date_value_2


# Функция - 4: Запись сводных данных в словарь
def write_dict_pg(pg_sheet_value, tmp_index_pg):
    check_value = 0
    shift_index = 87
    for key_value in range(8):
        date_value_1 = pg_sheet_value[tmp_index_pg][check_dict[key_value][0]].value
        date_value_2 = pg_sheet_value[tmp_index_pg][check_dict[key_value][1]].value
        date_value_1, date_value_2 = check_date_value(
            date_value_1, date_value_2)
        if key_value not in [3, 4]:
            write_dict_option_1(date_value_1, date_value_2, pg_sheet_value,
                                tmp_index_pg, check_value, key_value, shift_index)
        else:
            write_dict_option_2(date_value_1, date_value_2, pg_sheet_value,
                                tmp_index_pg, check_value, key_value, shift_index)
    return check_value, tmp_index_pg


# Функция - 5: Опциональная функция записи сводных данных в словарь (декомпозиция)
def write_dict_option_1(date_value_1, date_value_2, pg_sheet_value, tmp_index_pg, check_value, key_value, shift_index):
    check_flag = 0
    date_value = 0
    if date_value_1 is None:
        tmp_index_pg += 1
        check_value = 1
    if len(str(date_value_1)) == 10 and date_value_2 is None and check_value == 0:
        current_date = datetime.datetime.today()
        delta_date = check_dict[key_value][3]
        try:
            date_value = datetime.datetime.strptime(date_value_1, '%d.%m.%Y')
        except:
            check_flag = 1
        if check_flag == 0:
            if (current_date - date_value).days > delta_date:
                pg_sheet_value[tmp_index_pg][shift_index + key_value].value = str((
                    date_value + datetime.timedelta(days=delta_date)).strftime('%d.%m.%Y'))


# Функция - 6: Опциональная функция записи сводных данных в словарь (декомпозиция)
def write_dict_option_2(date_value_1, date_value_2, pg_sheet_value, tmp_index_pg, check_value, key_value, shift_index):
    check_flag = 0
    date_value = 0
    if date_value_1 is None:
        tmp_index_pg += 1
        check_value = 1
    if len(str(date_value_1)) == 10 and date_value_2 is None and check_value == 0:
        current_date = datetime.datetime.today()
        delta_date = check_dict[key_value][3]
        try:
            date_value = datetime.datetime.strptime(date_value_1, '%d.%m.%Y')
        except:
            check_flag = 1
        if check_flag == 0:
            if (current_date - date_value).days > delta_date[0] and pg_sheet_value[tmp_index_pg][2].value == 'CS':
                pg_sheet_value[tmp_index_pg][shift_index + key_value].value = str((
                    date_value + datetime.timedelta(days=delta_date[0])).strftime('%d.%m.%Y'))
            if (current_date - date_value).days > delta_date[1] and pg_sheet_value[tmp_index_pg][2].value == 'RT':
                pg_sheet_value[tmp_index_pg][shift_index + key_value].value = str((
                    date_value + datetime.timedelta(days=delta_date[1])).strftime('%d.%m.%Y'))
            if (current_date - date_value).days > delta_date[2] and pg_sheet_value[tmp_index_pg][2].value == 'GF':
                pg_sheet_value[tmp_index_pg][shift_index + key_value].value = str((
                    date_value + datetime.timedelta(days=delta_date[2])).strftime('%d.%m.%Y'))


# Функция - 7: Формирования сводного словаря просрочек
def create_dict_pg(pg_sheet_value, region_list, name_po):
    dict_pg = dict()
    dict_value_option = dict()
    for region_value in region_list:
        tmp_index_pg = 7
        while pg_sheet_value[tmp_index_pg][6].value == name_po:
            if pg_sheet_value[tmp_index_pg][0].value == region_value:
                if pg_sheet_value[tmp_index_pg][0].value not in dict_pg:
                    dict_value_option = dict()
                dict_value_option['кол-во БС'] = dict_value_option.get(
                    'кол-во БС', 0) + 1
                check_value, tmp_index_pg = write_dict_pg(
                    pg_sheet_value, tmp_index_pg)
                if check_value == 1:
                    continue
                dict_pg[pg_sheet_value[tmp_index_pg]
                        [0].value] = dict_value_option
                tmp_index_pg += 1
            else:
                tmp_index_pg += 1
    return dict_pg


# Функция - 8: Формирования поэтапного словаря просрочек
def create_dict_delay(pg_sheet_value, region_list):
    dict_delay = dict()
    for region_value in region_list:
        for tmp_index in range(87, 95):
            index_row = 7
            dict_delay_option = dict.fromkeys(range(13), 0)
            while pg_sheet_value[index_row][0].value:
                if pg_sheet_value[index_row][0].value == region_value and pg_sheet_value[index_row][tmp_index].value:
                    year_delay = int(
                        pg_sheet_value[index_row][tmp_index].value[6:10])
                    mount_delay = int(
                        pg_sheet_value[index_row][tmp_index].value[3:5])
                    if year_delay < int(date_year_value):
                        dict_delay_option[0] = dict_delay_option.get(0, 0) + 1
                    else:
                        dict_delay_option[mount_delay] = dict_delay_option.get(
                            mount_delay, 0) + 1
                index_row += 1
            dict_delay_option[13] = sum(dict_delay_option.values())
            dict_delay[(region_value, tmp_index)] = dict_delay_option
    return dict_delay


# Функция - 9: Запись данных из словарей (Сводный + по этапный) в реестр задержек
def write_delay_pivot(dict_pg, dict_delay, index_row_delay, name_po):
    delay_book_value = openpyxl.open(tmp_dir + '\Reestr_delay.xlsx')
    delay_sheet_value = delay_book_value['реестр']
    for key_1, value_1 in dict_pg.items():
        tmp_index = 3
        delay_sheet_value[index_row_delay][0].value = name_po
        delay_sheet_value[index_row_delay][1].value = key_1
        delay_sheet_value[index_row_delay][2].value = str(
            week_value) + date_year_value
        for tmp_value in value_1.values():
            delay_sheet_value[index_row_delay][tmp_index].value = tmp_value
            tmp_index += 1
        for key_2, value_2 in dict_delay.items():
            if key_2[0] == key_1:
                for tmp_value in value_2.values():
                    delay_sheet_value[index_row_delay][tmp_index].value = tmp_value
                    tmp_index += 1
        index_row_delay += 1
    delay_book_value.save(tmp_dir + '\Reestr_delay.xlsx')
    delay_book_value.close()
    return index_row_delay


# Функция - 10: Архивирования реестра задержек
def create_zip():
    value_password = bytes(
        str(datetime.datetime.today().strftime("%d%m")), "utf-8")
    for root_value, dirs_value, files_value in os.walk(set_dir):
        for file_name in files_value:
            with ZipFile(f'{set_dir}\{file_name[:len(file_name) - 5]}' + '.zip', 'w') as zip_value:
                zip_value.write(
                    f'{set_dir}\{file_name}', arcname=f'{file_name}', pwd=b"%s" % (value_password))
            zip_value.close()


# Функция - 11: Функция по созданию ограничений на ввод данных
def create_contain(pg_sheet_value, file_name):
    dv_value_list = DataValidation(
        type='list', formula1='Факторы!$A$1:$A$14', allow_blank=True)
    dv_value_list.error = error_value_list
    dv_value_list.errorTitle = 'Внимание, указаны не корректные данные'
    dv_value_date = DataValidation(type='date')
    dv_value_date.error = error_value_date
    dv_value_date.errorTitle = 'Внимание, указаны не корректные данные'
    pg_sheet_value.add_data_validation(dv_value_list)
    pg_sheet_value.add_data_validation(dv_value_date)
    row = len(pandas.read_excel(set_dir + '\\' + file_name,
              sheet_name='ПГ', header=5).values.tolist()) + 6
    for range_value in ('J7:M{row_index}', 'T7:AA{row_index}', 'AD7:AY{row_index}'):
        dv_value_date.add(range_value.format(row_index=row))
    dv_value_list.add(f'BA7:BA{row}')


# Функция - 12: Основная функция по формирования словарей (Сводный + по этапный)
def check_pg():
    index_row_delay = rows_reestrs()
    for set_r_value, set_d_value, set_f_value in os.walk(set_dir):
        for file_name in set_f_value:
            if file_name.endswith('.xlsx'):
                name_po = str(file_name)[23:-14]
                region_list = create_region_list(file_name)
                pg_book_value = openpyxl.open(set_dir + '\\' + file_name)
                pg_sheet_value = pg_book_value['ПГ']
                dict_pg = create_dict_pg(pg_sheet_value, region_list, name_po)
                dict_delay = create_dict_delay(pg_sheet_value, region_list)
                index_row_delay = write_delay_pivot(
                    dict_pg, dict_delay, index_row_delay, name_po)
                create_contain(pg_sheet_value, file_name)
                pg_book_value.save(set_dir + '\\' + file_name)
                pg_book_value.close()


# 3 этап: Запуск основной функции/программы
def main_xl_3():
    art_print_start(name_project='Check PG')
    data_start = datetime.datetime.now()
    check_pg()
    create_zip()
    art_print_end(data_start)
    successfully_send(data_start, name_project='Check PG')


if __name__ == '__main__':
    main_xl_3()
