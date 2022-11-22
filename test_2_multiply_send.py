from outlook_templates import successfully_send, send_index_list, html_value_1, art_print_start, art_print_end, \
    function_sort, html_value_2_1, html_value_2_2, html_value_2_3, html_value_2_4, html_value_2_5
from pathlib import Path
import win32com.client
import pythoncom
import tabulate
import datetime
import openpyxl
import pandas
import os

# Скрипт по автоматической рассылке в адрес руководства результатов/итогов работы с ПО, в части получения, анализа ПГ
# Разработчик: Ilya Turkhanovich
# Версия: 1.2 (Используется адаптивное формирование текста письма для рассылки в адрес руководства)
# Статус: Автоматический запуск, тестирование
# Расписание запуска: По понедельника каждую четную неделю, в 09:00 по Нск

# ----------------------------
# Этапы (оглавление):
# 1 этап: Подготовка исходных данных, объявление необходимых переменных, списков/массивов
# 2 этап: Функции автоматической рассылке (Общее кол-во: 6 функций)
# 3 этап: Запуск основной функции/программы

# 1 этап: Глобальные переменные
dir_value = str(Path.cwd())
tmp_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_templates'
back_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_backup'
set_dir = dir_value[:dir_value.rfind('\\')] + r'\outlook_set'
tmp_date_year = datetime.datetime.today().strftime("%Y")
date_reestr = str(datetime.datetime.now().strftime('%d%m%Y'))
week_value = 'W' + str(datetime.datetime.now().isocalendar()[1] - 2)


# 2 этап: Функции
# Функция - 1: Формирование реестра в формате .xlsx по текущей неделе, а также свода по ПО для рассылки в адрес руководства
def info_reestr():
    df_value = pandas.read_excel(tmp_dir + r'\Reestr.xlsx', sheet_name='реестр').values.tolist()
    po_list_not = list()
    reestr_list = list()
    po_list_error = list()
    po_list_incorrect = list()
    po_list = set()
    po_list_not.append(('Подрядная организация', 'Дата запроса'))
    po_list_incorrect.append(('Подрядная организация', 'Дата запроса'))
    po_list_error.append(('Подрядная организация', 'Дата запроса'))
    book_value = openpyxl.open(back_dir + r'\Reestr_cypris.xlsx')
    sheet_value = book_value.active
    row_index = 2
    for row_value in df_value:
        if row_value[1] == week_value:
            po_list.add(row_value[0])
            dict_reesrt = {0: 0, 2: 1, 3: 2, 4: 3, 6: 4, 7: 5, 8: 6, 9: 7}
            for col_index in dict_reesrt:
                if pandas.isna(row_value[col_index]):
                    sheet_value[row_index][dict_reesrt[col_index]].value = ''
                else:
                    sheet_value[row_index][dict_reesrt[col_index]].value = row_value[col_index]
                sheet_value[row_index][8].value = ', '.join(detect_cypris(po_name=row_value[0]))
            row_index += 1
        if row_value[1] == week_value and row_value[9] == 'да':
            status = row_value[4]
            name_po = row_value[0]
            dict_value = dict()
            for tmp_index in df_value:
                if tmp_index[0] == name_po:
                    dict_value[int(tmp_index[1][1:])] = (tmp_index[2], tmp_index[4])
            for tmp_key, tmp_value in sorted(dict_value.items()):
                date_value = tmp_value[0]
                if tmp_value[1] == status and status == 'ответ не получен':
                    po_list_not.append((name_po, date_value))
                    break
                if tmp_value[1] == status and status == 'не по шаблону':
                    po_list_incorrect.append((name_po, date_value))
                    break
                if tmp_value[1] == status and status == 'с ошибками':
                    po_list_error.append((name_po, date_value))
                    break
        if row_value[1] == week_value and pandas.isna(row_value[9]) and (row_value[8] == 'нет' or
                                                                         pandas.isna(row_value[8])):
            reestr_list.append(row_value[4])
    book_value.save(back_dir + r'\Reestr_cypris_' + date_reestr + '.xlsx')
    book_value.close()
    return reestr_list, po_list, po_list_not, po_list_incorrect, po_list_error


# Функция - 2: Определение подраделения в котором работает соответсвующий ПО
def detect_cypris(po_name):
    cypris_set = set()
    for set_r_value, set_d_value, set_f_value in os.walk(set_dir):
        for file_name in set_f_value:
            name_po_set = str(file_name)[23:-14]
            if po_name == name_po_set:
                df_value = pandas.read_excel(set_dir + '\\' + file_name, sheet_name='ПГ', header=5).values.tolist()
                for row_value in df_value:
                    cypris_set.add(row_value[53])
    return sorted(tuple(cypris_set))


# Функция - 3: Формирование реестра просрочек в формате словаря по текущей неделе
def create_dict_delay(po_list):
    df_value = pandas.read_excel(tmp_dir + r'\Reestr_delay.xlsx', sheet_name='реестр').values.tolist()
    week_year_value = str(week_value + tmp_date_year)
    dict_delay = dict()
    for name_po in sorted(po_list):
        dict_delay[name_po] = dict.fromkeys(send_index_list, 0)
        for delay_row in df_value:
            if delay_row[0] == name_po and delay_row[2] == week_year_value:
                for tmp_index in send_index_list:
                    dict_delay[name_po][tmp_index] = dict_delay[name_po].get(tmp_index, 0) + delay_row[tmp_index]
    list_delay = list()
    for key, value in sorted(dict_delay.items()):
        tmp_index, count_bs = 0, 0
        smr_max, doc_max, check_max = 0, 0, 0
        for tmp_value in value.values():
            if tmp_index == 0:
                count_bs = tmp_value
            if tmp_index > 0 and tmp_index < 4 and tmp_value > doc_max:
                doc_max = tmp_value
            if tmp_index > 3 and tmp_index < 7 and tmp_value > smr_max:
                smr_max = tmp_value
            if tmp_index > 6 and tmp_index < 9 and tmp_value > check_max:
                check_max = tmp_value
            tmp_index += 1
        mid_index = doc_max / count_bs * 0.25 + smr_max / count_bs * 0.7 + check_max / count_bs * 0.05
        list_delay_option = list((key, count_bs, smr_max/count_bs, mid_index, doc_max/count_bs, check_max/count_bs))
        list_delay.append(tuple(list_delay_option))
        list_delay_option.clear()
    dict_pivot_delay_1 = dict(zip(range(1, len(list_delay) + 1),
                                  sorted(list_delay, reverse=True, key=function_sort(2))))
    dict_pivot_delay_2 = dict(zip(range(1, len(list_delay) + 1),
                                  sorted(list_delay, reverse=True, key=function_sort(3))))
    tmp_rang = 0
    for key_1, value_1 in sorted(dict_delay.items()):
        for key_2, value_2 in dict_pivot_delay_1.items():
            if key_1 == value_2[0]:
                dict_delay[key_1]['smr_max'] = round(value_2[2] * 100)
                dict_delay[key_1]['doc_max'] = round(value_2[4] * 100)
                dict_delay[key_1]['check_max'] = round(value_2[5] * 100)
                dict_delay[key_1]['mid_index'] = round(value_2[3] * 100)
                tmp_rang = key_2
                break
        for key_2, value_2 in dict_pivot_delay_2.items():
            if key_1 == value_2[0]:
                dict_delay[key_1]['mid_rang'] = key_2 * tmp_rang
                break
    info_delay(dict_delay)


# Функция - 4: Запись словаря просрочек в реестр просрочек формата .csv
def info_delay(dict_delay):
    book_value = openpyxl.open(back_dir + r'\Reestr_delay_cypris.xlsx')
    sheet_value = book_value.active
    index_value = 0
    row_index = 3
    for key, value in sorted(dict_delay.items()):
        index_value += 1
        sheet_value[row_index][0].value = index_value
        sheet_value[row_index][1].value = key
        col_index = 2
        for tmp_value in value.values():
            if tmp_value == 0:
                sheet_value[row_index][col_index].value = ''
            else:
                sheet_value[row_index][col_index].value = tmp_value
            col_index += 1
        row_index += 1
    book_value.save(back_dir + r'\Reestr_delay_cypris_' + date_reestr + '.xlsx')
    book_value.close()


# Функция - 5: Функция рассылка письма в адрес руководства
def send_mail():
    outlook_value = win32com.client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    send_account = None
    for account in outlook_value.Session.Accounts:
        if account.DisplayName == "******@*******.****":
            send_account = account
            break
    message_value = outlook_value.CreateItem(0)
    message_value._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    message_value.Subject = f"План-графики по объектам сети радиодоступа, итоги за {week_value}"
    create_html(message_value)
    message_value.Attachments.Add(back_dir + r'\Reestr_cypris_' + date_reestr + '.xlsx')
    message_value.Attachments.Add(back_dir + r'\Reestr_delay_cypris_' + date_reestr + '.xlsx')
    message_value.To = """*****@****.****; *****@****.****; *****@****.****; *****@****.****;
    *****@****.****; *****@****.****; *****@****.****;"""
    message_value.Send()
    os.remove(back_dir + r'\Reestr_cypris_' + date_reestr + '.xlsx')
    os.remove(back_dir + r'\Reestr_delay_cypris_' + date_reestr + '.xlsx')


# Функция - 6: Формирования адаптивного текста письма для рассылки в адрес руководства
def create_html(message_value):
    reestr_list, po_list, po_list_not, po_list_incorrect, po_list_error = info_reestr()
    create_dict_delay(po_list)
    count_not = len(po_list_not) - 1
    count_incorrect = len(po_list_incorrect) - 1
    count_error = len(po_list_error) - 1
    message_value.HTMLBody = html_value_1.format(week=week_value, count_total=len(po_list),
                                                 count_correct=reestr_list.count('по шаблону'),
                                                 count_no_temp_1=reestr_list.count('не по шаблону'),
                                                 count_error_pg_1=reestr_list.count('с ошибками'),
                                                 count_no_pg_1=reestr_list.count('ответ не получен'))
    if (count_error + count_not + count_incorrect) != 0:
        message_value.HTMLBody += html_value_2_1
    if count_incorrect != 0:
        message_value.HTMLBody += html_value_2_2.format(count_no_temp_2=count_incorrect,
                                                        table_1=tabulate.tabulate(po_list_incorrect, headers="firstrow",
                                                                                  tablefmt="html", stralign='left'))
    if count_error != 0:
        message_value.HTMLBody += html_value_2_3.format(count_error_pg_2=count_error,
                                                        table_2=tabulate.tabulate(po_list_error, headers="firstrow",
                                                                                  tablefmt="html", stralign='left'))
    if count_not != 0:
        message_value.HTMLBody += html_value_2_4.format(count_no_pg_2=count_not,
                                                        table_3=tabulate.tabulate(po_list_not, headers="firstrow",
                                                                                  tablefmt="html", stralign='left'))

    message_value.HTMLBody += html_value_2_5.format(week=week_value)


# 3 этап: Запуск основной функции/программы
def main_send_2():
    art_print_start(name_project='Send mail')
    data_start = datetime.datetime.now()
    send_mail()
    art_print_end(data_start)
    successfully_send(data_start, name_project='Send mail')


if __name__ == '__main__':
    main_send_2()
