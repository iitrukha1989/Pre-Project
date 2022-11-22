from outlook_templates import art_print_start, art_print_end, successfully_send, dict_cypris, \
    html_value_5_1, html_value_5_2, html_value_5_3, html_value_5_4, html_value_5_5, html_value_5_6, html_value_5_7, \
    html_value_5_8, html_value_5_9, html_value_5_10
from test_3_xl import create_dict_pg, create_region_list
import win32com.client
from win32com.client import constants
from win32com.client.gencache import EnsureDispatch
from zipencrypt import ZipFile
import pythoncom
import datetime
import openpyxl
import pathlib
import pandas
import os

# Скрипт по сверке полученных от ПО план - графиков (далее по тексту ПГ) на предмет валидности по шаблону
# Разработчик: Ilya Trukhanovich
# Версия: 1.3 (Используется адаптивное формирования текста письма повторного запроса, анализ обновленного шаблона ПГ)
# Статус: Автоматический запуск, тестирование
# Расписание запуска: По четвергам и субботам каждую не четную неделю, после выполнения test_1_multiply_get.py

# ----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции по проверке ПГ (Общее кол-во: 19 функций)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(pathlib.Path.cwd())
tmp_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_templates'
set_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_set'
get_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_get'
week_value = 'W' + str(datetime.datetime.now().isocalendar()[1] - 1)
date_value = datetime.datetime.now().strftime('%d%m%Y')


# Функция - 1: Первая итерация проверки полученных ПГ от ПО, перенос информации полученной от ПО в актуальный ПГ
def check_get_1():
    for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
        for tmp_number in create_get_list():
            for get_file in get_f_value:
                if tmp_number == int(str(get_file)[:str(get_file).find('_')]):
                    name_po_get, sender_date, pg_tuple = convert_df_tuple(get_file)
                    os.remove(f"{get_dir}\{get_file}")
                    check_value = 0
                    for get_r_value_1, get_d_value_1, get_f_value_1 in os.walk(get_dir):
                        for get_file_1 in get_f_value_1:
                            if get_file_1 == str(r'План-график БС' + name_po_get + '_' + date_value + '.xlsx'):
                                book_value = openpyxl.open(get_dir + '\\' + get_file_1)
                                check_get_option(pg_tuple, book_value, name_po_get, sender_date)
                                check_value = 1
                    if check_value == 0:
                        for set_r_value, set_d_value, set_f_value in os.walk(set_dir):
                            for set_file in set_f_value:
                                name_po_set = str(set_file)[23:-14]
                                if name_po_get == name_po_set:
                                    book_value = openpyxl.open(set_dir + '\\' + set_file)
                                    check_get_option(pg_tuple, book_value, name_po_get, sender_date)


# Функция - 2: Функция получения порядковых номеров ПГ
def create_get_list():
    for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
        get_list = list()
        for file_name in sorted(get_f_value):
            get_list.append(int(str(file_name)[:str(file_name).find('_')]))
    return sorted(get_list, reverse=True)


# Функция - 3: Опционная функция переноса информации полученной от ПО в актуальный ПГ
def check_get_option(pg_tuple, book_value, name_po_get, sender_date):
    sheet_value = book_value.active
    row_index = 7
    pl_set = set()
    while sheet_value[row_index][0].value:
        pl_set.add((sheet_value[row_index][1].value, sheet_value[row_index][2].value, sheet_value[row_index][7].value))
        for row_value in pg_tuple:
            if row_value[1] == sheet_value[row_index][1].value and \
                    row_value[2] == sheet_value[row_index][2].value and \
                    row_value[7] == sheet_value[row_index][7].value:
                for col_index in range(9, 50, 2):
                    if pandas.isna(row_value[col_index]):
                        sheet_value[row_index][col_index].value = ''
                    elif row_value[col_index] in ['не треб', 'не треб.', 'не требуется']:
                        sheet_value[row_index][col_index].value = 'не треб'
                    elif isinstance(row_value[col_index], datetime.datetime):
                        sheet_value[row_index][col_index].value = str(row_value[col_index].strftime('%d.%m.%Y'))
                for col_index in range(10, 51, 2):
                    if pandas.isna(row_value[col_index]) is False and \
                            (sheet_value[row_index][col_index].value is None or
                             sheet_value[row_index][col_index].value == ''):
                        if row_value[col_index] in ['не треб', 'не треб.', 'не требуется']:
                            sheet_value[row_index][col_index].value = 'не треб'
                        if isinstance(row_value[col_index], datetime.datetime):
                            sheet_value[row_index][col_index].value = str(row_value[col_index].strftime('%d.%m.%Y'))
                if pandas.isna(row_value[51]):
                    sheet_value[row_index][51].value = ''
                else:
                    sheet_value[row_index][51].value = row_value[51]
                if pandas.isna(row_value[52]):
                    sheet_value[row_index][52].value = ''
                else:
                    sheet_value[row_index][52].value = row_value[52]
        row_index += 1
    for row_value in pg_tuple:
        if (row_value[1], row_value[2], row_value[7]) not in pl_set and row_value[6] == name_po_get:
            for col_index in range(53):
                sheet_value[row_index][col_index].value = row_value[col_index]
            sheet_value[row_index][77].value = 'не отражен в план-графике'
            sheet_value[row_index][53].value = dict_cypris[row_value[0]]
            row_index += 1
    sheet_value[1][78].value = datetime.datetime.strptime(sender_date[:2] + '.' +
                                                          sender_date[2:4] + '.' +
                                                          sender_date[4:], '%d.%m.%Y')
    book_value.save(get_dir + str(r'\План-график БС ПАО МТС_' + name_po_get + '_' + date_value + '.xlsx'))
    book_value.close()


# Функция - 4: Преобразование DataFrame в Tuple (в кортеж, для  надежности и скорости обращения к ПГ)
def convert_df_tuple(get_file):
    sheet_name = openpyxl.open(get_dir + '\\' + get_file).sheetnames[0]
    name_po_get = openpyxl.open(get_dir + '\\' + get_file).active[7][6].value
    sender_date = get_file[get_file.find('_') + 1:get_file.find('_') + 5] + datetime.datetime.now().strftime("%Y")
    df_value = pandas.read_excel(get_dir + '\\' + get_file, sheet_name=sheet_name, header=5).values.tolist()
    list_option = list()
    for tmp_row in df_value:
        if pandas.isna(tmp_row[0]) or pandas.isna(tmp_row[1]):
            break
        list_option.append(tuple(tmp_row))
    return name_po_get, sender_date, tuple(list_option)


# Функция - 5: Вторая итерация проверки полученных ПГ от ПО, перезапись формул расчета ошибок
def check_get_2():
    excel_api = win32com.client.Dispatch("Excel.Application")
    excel_api.Visible = False
    tmp_win32 = win32com.client.constants
    for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
        for file_name in get_f_value:
            len_df = int(len(pandas.read_excel(get_dir + '\\' + file_name,
                                               sheet_name='ПГ', header=5).values.tolist()) + 6)
            if len_df > 7:
                book_value = excel_api.Workbooks.Open(get_dir + '\\' + file_name)
                sheet_value = book_value.ActiveSheet
                xl = EnsureDispatch('Excel.Application')
                xl.Range("BC7:BY7").Select()
                xl.Selection.AutoFill(Destination=sheet_value.Range(f"BC7:BY{len_df}"), Type=constants.xlFillDefault)
                book_value.Close(True)


# Функция - 6: Третья итерация проверки полученных ПГ от ПО, запись результов проверки полученных ПГ в реестр
def check_get_3():
    for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
        for file_name in get_f_value:
            name_po = str(file_name)[23:-14]
            excel_api = win32com.client.Dispatch("Excel.Application")
            book_value = excel_api.Workbooks.Open(get_dir + '\\' + file_name)
            xl = EnsureDispatch('Excel.Application')
            error_value_1 = round(xl.Range("CC3").Value * 100, 0)
            error_value_2 = int(xl.Range("CD3").Value)
            sender_date = str(xl.Range("CA1").Value)
            book_value.Close(False)
            if error_value_1 >= 28:
                log_reestr(name_po, error_value_1, error_value_2, sender_date, mode_send=1)
            else:
                log_reestr(name_po, error_value_1, error_value_2, sender_date, mode_send=2)


# Функция - 7: Опционная функция записи в реестр
def log_reestr(name_po, error_value_1, error_value_2, sender_date, mode_send):
    book_value = openpyxl.load_workbook(tmp_dir + r'\Reestr.xlsx')
    sheet_value = book_value.active
    index_row = 1
    while sheet_value[index_row][0].value:
        index_row += 1
    for tmp_index in range(1, index_row):
        if sheet_value[tmp_index][0].value == name_po and \
                sheet_value[tmp_index][1].value == week_value and \
                sheet_value[tmp_index][4].value is None:
            if mode_send == 1:
                sheet_value[tmp_index][3].value = sender_date[8:10] + '.' + sender_date[5:7] + '.' + sender_date[:4]
                sheet_value[tmp_index][4].value = 'с ошибками'
                sheet_value[tmp_index][6].value = error_value_1
                sheet_value[tmp_index][7].value = error_value_2
                if datetime.datetime.isoweekday(datetime.datetime.now()) == 4:
                    sheet_value[tmp_index][5].value = datetime.date.today().strftime("%d.%m.%Y")
                    sheet_value[tmp_index][8].value = 'да'
                    sheet_value.insert_rows(tmp_index + 1, 1)
                    sheet_value[tmp_index + 1][0].value = name_po
                    sheet_value[tmp_index + 1][1].value = sheet_value[tmp_index][1].value
                    sheet_value[tmp_index + 1][2].value = datetime.date.today().strftime("%d.%m.%Y")
                break
            if mode_send == 2:
                sheet_value[tmp_index][3].value = sender_date[8:10] + '.' + sender_date[5:7] + '.' + sender_date[:4]
                sheet_value[tmp_index][4].value = 'по шаблону'
                sheet_value[tmp_index][6].value = error_value_1
                sheet_value[tmp_index][7].value = error_value_2
                sheet_value[tmp_index][8].value = 'нет'
                break
    book_value.save(tmp_dir + r'\Reestr.xlsx')
    book_value.close()


# Функция - 8: Четвертая итерация проверки полученных ПГ от ПО, запись не полученных ПГ в реестр
def check_get_4():
    book_value = openpyxl.open(tmp_dir + r'\Reestr.xlsx')
    sheet_value = book_value.active
    index_row, po_reestr_set = create_not_po()
    for po_name in po_reestr_set:
        for tmp_index in range(1, index_row + 1):
            if sheet_value[tmp_index][0].value == po_name and \
                    sheet_value[tmp_index][1].value == week_value and sheet_value[tmp_index][4].value is None:
                sheet_value[tmp_index][4].value = 'ответ не получен'
                if datetime.datetime.isoweekday(datetime.datetime.now()) == 4:
                    sheet_value[tmp_index][5].value = datetime.date.today().strftime("%d.%m.%Y")
                    sheet_value[tmp_index][8].value = 'да'
                    sheet_value.insert_rows(tmp_index + 1, 1)
                    sheet_value[tmp_index + 1][0].value = sheet_value[tmp_index][0].value
                    sheet_value[tmp_index + 1][1].value = sheet_value[tmp_index][1].value
                    sheet_value[tmp_index + 1][2].value = datetime.date.today().strftime("%d.%m.%Y")
                    index_row += 1
                break
    book_value.save(tmp_dir + r'\Reestr.xlsx')
    book_value.close()


# Функция - 9: Опционная функция определения списка(множества) ПО по которым не получен ответ
def create_not_po():
    df_value = pandas.read_excel(tmp_dir + r'\Reestr.xlsx', sheet_name='реестр').values.tolist()
    index_row = 1
    po_reestr_set_1 = set()
    po_reestr_set_2 = set()
    for tmp_row in tuple(df_value):
        index_row += 1
        if datetime.datetime.isoweekday(datetime.datetime.now()) == 4:
            if tmp_row[1] == week_value:
                po_reestr_set_1.add(tmp_row[0])
            if tmp_row[1] == week_value and pandas.isna(tmp_row[4]) is False:
                po_reestr_set_2.add(tmp_row[0])
        else:
            if tmp_row[1] == week_value and pandas.isna(tmp_row[4]):
                po_reestr_set_1.add(tmp_row[0])
    return index_row, po_reestr_set_1 - po_reestr_set_2


# Функция - 10: Функция повторного пересчета просрочек
def create_delay():
    if datetime.datetime.isoweekday(datetime.datetime.now()) == 4:
        for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
            for file_name in get_f_value:
                if file_name.endswith('.xslx'):
                    name_po = str(file_name)[23:-14]
                    region_list = create_region_list(file_name)
                    pg_book_value = openpyxl.open(set_dir + '\\' + file_name)
                    pg_sheet_value = pg_book_value['ПГ']
                    create_dict_pg(pg_sheet_value, region_list, name_po)


# Функция - 11: Опционная функция архивирование ПГ для отправки запросов в адрес ПО (запуск: только по четверг)
def create_zip():
    if datetime.datetime.isoweekday(datetime.datetime.now()) == 4:
        value_password = bytes(str(datetime.datetime.today().strftime("%d%m")), "utf-8")
        for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
            for file_name in get_f_value:
                if file_name.endswith('.xlsx'):
                    with ZipFile(f'{get_dir}\{file_name[:len(file_name) - 5]}' + '.zip', 'w') as zip_value:
                        zip_value.write(f'{get_dir}\{file_name}', arcname=f'{file_name}', pwd=b"%s" % (value_password))
                    zip_value.close()
        for set_r_value, set_d_value, set_f_value in os.walk(set_dir):
            for file_name in set_f_value:
                if file_name.endswith('.xlsx'):
                    with ZipFile(f'{set_dir}\{file_name[:len(file_name) - 5]}' + '.zip', 'w') as zip_value:
                        zip_value.write(f'{set_dir}\{file_name}', arcname=f'{file_name}', pwd=b"%s" % (value_password))
                    zip_value.close()


# Функция - 12: Анализ реестра ПГ, на предмет необходимости формирования повторного запроса (запуск: только по четверг)
def analiz_reestr():
    df_value_reestr = pandas.read_excel(tmp_dir + r'\Reestr.xlsx', sheet_name='реестр').values.tolist()
    df_value_contact = pandas.read_excel(tmp_dir + r'\name_inn_po.xlsx', sheet_name='name_inn_po').values.tolist()
    if datetime.datetime.isoweekday(datetime.datetime.now()) == 4:
        for index_row_reestr in df_value_reestr:
            if index_row_reestr[1] == week_value and index_row_reestr[8] == 'да':
                name_po = index_row_reestr[0]
                contact_po = ''
                status_pg = index_row_reestr[4]
                for index_row_contact in df_value_contact:
                    if index_row_reestr[0] == index_row_contact[0]:
                        contact_po = index_row_contact[2]
                send_mail(name_po, contact_po, status_pg)
    else:
        analiz_reesrt_option(df_value_reestr)


# Функция - 13: Опционая функция отправки писем в адрес ПО (запуск: только по четверг)
def send_mail(name_po, contact_po, status_pg):
    po_dict = create_send_delay(name_po, status_pg)
    res_status = correct_status(status_pg)
    outlook_api = win32com.client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    send_account = None
    for account in outlook_api.Session.Accounts:
        if account.DisplayName == "cupris-vostok@mts.ru":
            send_account = account
            break
    message_value = outlook_api.CreateItem(0)
    message_value._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    message_value.Subject = f"План-график БС ПАО МТС_{name_po}_(повторный)"
    create_html(message_value, po_dict, res_status)
    if status_pg == 'с ошибками':
        message_value.Attachments.Add(get_dir + f'\План-график БС ПАО МТС_{name_po}_{date_value}.zip')
    else:
        date_value_1 = (datetime.datetime.now() - datetime.timedelta(days=6)).strftime("%d%m%Y")
        message_value.Attachments.Add(set_dir + f'\План-график БС ПАО МТС_{name_po}_{date_value_1}.zip')
    message_value.To = contact_po
    message_value.CC = '******@****.***; ******@****.***; ******@****.***;'
    message_value.Send()


# Функция - 14: Опционная функция изменения формата ответа
def correct_status(status_pg):
    res_status = ''
    if status_pg == 'ответ не получен':
        res_status = 'не получен'
    elif status_pg == 'с ошибками':
        res_status = 'получен ' + status_pg + ' и требуется его дозаполнение в части фактических или планируемых дат'
    elif status_pg == 'не по шаблону':
        res_status = 'получен ' + status_pg + ' формат файла не .xlsx/файл заархивирован/файл поврежден, не открывается'
    return res_status


# Функция - 15: Опционная функция создания словарей задержек для формирования фактуры в рассылку
def create_send_delay(name_po, status_pg):
    po_dict = {87: list(), 90: list(), 91: list(), 92: list(), 93: list(), 94: list()}
    if status_pg == 'с ошибками':
        df_value = pandas.read_excel(get_dir + f'\План-график БС ПАО МТС_{name_po}_{date_value}.xlsx',
                                     sheet_name='ПГ', header=5).values.tolist()
    else:
        date_value_1 = (datetime.datetime.now() - datetime.timedelta(days=6)).strftime("%d%m%Y")
        df_value = pandas.read_excel(set_dir + f'\План-график БС ПАО МТС_{name_po}_{date_value_1}.xlsx',
                                     sheet_name='ПГ', header=5).values.tolist()
    for row_value in df_value:
        for tmp_index in [87, 90, 91, 92, 93, 94]:
            if pandas.isna(row_value[tmp_index]) is False:
                po_dict[tmp_index].append(row_value[1])
    return po_dict


# Функция - 16: Формирования адаптивного текста письма для рассылки в адрес ПО
def create_html(message_value, po_dict, res_status):
    check_sum = 0
    message_value.HTMLBody = html_value_5_1.format(date_value_1=(datetime.datetime.now() -
                                                                 datetime.timedelta(days=6)).strftime("%d.%m.%Y"),
                                                   status=res_status)
    for tmp_index in [87, 90, 91, 92, 93, 94]:
        check_sum += int(len(po_dict[tmp_index]))
    if check_sum != 0:
        message_value.HTMLBody += html_value_5_2
    if len(po_dict[87]) != 0:
        message_value.HTMLBody += html_value_5_3.format(count_to=len(po_dict[87]), list_to=', '.join(po_dict[87]))
    if len(po_dict[90]) != 0:
        message_value.HTMLBody += html_value_5_4.format(count_smr=len(po_dict[90]), list_smr=', '.join(po_dict[90]))
    if len(po_dict[91]) != 0:
        message_value.HTMLBody += html_value_5_5.format(count_smrg=len(po_dict[91]), list_smrg=', '.join(po_dict[91]))
    if len(po_dict[92]) != 0:
        message_value.HTMLBody += html_value_5_6.format(count_check=len(po_dict[92]), list_check=', '.join(po_dict[92]))
    if len(po_dict[93]) != 0:
        message_value.HTMLBody += html_value_5_7.format(count_ks=len(po_dict[93]), list_ks=', '.join(po_dict[93]))
    if len(po_dict[94]) != 0:
        message_value.HTMLBody += html_value_5_8.format(count_aop=len(po_dict[94]), list_aop=', '.join(po_dict[94]))
    if check_sum != 0:
        message_value.HTMLBody += html_value_5_9
    message_value.HTMLBody += html_value_5_10.format(date_value_2=(datetime.datetime.now() +
                                                                   datetime.timedelta(days=1)).strftime("%d.%m.%Y"))


# Функция - 17: Опционный анализ реестра на предмет поиска ПО по которым требуется АКС
def analiz_reesrt_option(df_value_reestr):
    book_value = openpyxl.open(tmp_dir + r'\Reestr.xlsx')
    sheet_value = book_value.active
    po_set = set()
    status_set = set()
    pr_week = week_value[0] + str(int(week_value[1:]) - 2)
    for row_value in df_value_reestr:
        if row_value[1] == week_value:
            po_set.add(row_value[0])
    for po_value in po_set:
        for row_value in df_value_reestr:
            if row_value[0] == po_value and (row_value[1] == pr_week or row_value[1] == week_value):
                status_set.add(row_value[4])
        if len(status_set) == 1 and status_set.pop() != 'по шаблону':
            row_index = 2
            while sheet_value[row_index][0].value:
                if sheet_value[row_index][0].value == po_value and sheet_value[row_index][1].value == week_value and \
                        (sheet_value[row_index][8].value is None or sheet_value[row_index][8].value == ''):
                    sheet_value[row_index][9].value = 'да'
                row_index += 1
        status_set.clear()
    book_value.save(tmp_dir + r'\Reestr.xlsx')
    book_value.close()


# Функция - 18: Функция очистки папки outlook_set/get от ПГ
def clear_dir():
    if datetime.datetime.isoweekday(datetime.datetime.now()) == 4:
        for get_r_value, get_d_value, get_f_value in os.walk(get_dir):
            for file_name in get_f_value:
                os.remove(f"{get_dir}\{file_name}")
    for set_r_value, set_d_value, set_f_value in os.walk(set_dir):
        for file_name in set_f_value:
            if file_name.endswith('.zip'):
                os.remove(f"{set_dir}\{file_name}")


# Функция - 19: Основная функция анализа полученных ПГ
def check_pg():
    check_get_1()
    check_get_2()
    check_get_3()
    check_get_4()
    create_delay()
    create_zip()
    analiz_reestr()
    clear_dir()


# 3 этап: Запуск основной функции/программы
def main_xl_2():
    art_print_start(name_project='Check PG')
    date_start = datetime.datetime.now()
    check_pg()
    art_print_end(date_start)
    successfully_send(date_start, name_project='Check PG')


if __name__ == '__main__':
    main_xl_2()
