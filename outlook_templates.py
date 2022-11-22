import datetime
import win32com.client as client

#  Набор констант: функции, строки, кортежей и словарей используемых в скриптах
#  Разработчик: Ilya Trukhanovich

# ----------------------------
#  Этапы (оглавление):
#  1 этап: Константные функции и методы
#  2 этап: Строковые константы
#  3 этап: Кортежные константы
#  4 этап: Словарные константы


#  Константные функции и методы
#  Функция - 1: Информация о запуске программы/скрипта
def art_print_start(name_project):
    print(f'[+] Project: {name_project}')
    print('[+] Developer: Ilya Trukhanovich')
    print('[+] Status project: testing/update')
    print('[+] Start script on time:',
          datetime.datetime.today().strftime('%H:%M'))
    print('[+] proccesing ...')


#  Функция - 2: Информация о завершении работы программы/скрипта, с указанием времени его работы
def art_print_end(data_start):
    print('[+] End script on time:',
          datetime.datetime.today().strftime('%H:%M'))
    data_end = datetime.datetime.now()
    print('[+] Execution time script:',
          round(float((data_end - data_start).total_seconds()), 1), 's.')
    print("[+] Powered by ******")


#  Функция - 3: Почтовое оповещение об успешном завершении работы программы/скрипта (удалено из кода/скорректировано в коде)
def successfully_send(data_start, name_project):
    text = f"""Project: {name_project} - completed successfully.
    
    Start script on time:', {data_start.strftime('%H:%M')}
    End script on time:', {datetime.datetime.today().strftime('%H:%M')}
    Execution time script:', {round(float((datetime.datetime.now() - data_start).total_seconds()), 1)}s."""
    tmp_outlook_app = client.Dispatch("Outlook.Application")
    send_account = None
    for account in tmp_outlook_app.Session.Accounts:
        if account.DisplayName == "******@*****.***":
            send_account = account
            break
    tmp_message = tmp_outlook_app.CreateItem(0)
    tmp_message._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    tmp_message.Subject = "План-графики по объектам сети радиодоступа, скрипт по формированию."
    tmp_message.To = "******@*****.****"
    tmp_message.Body = text
    tmp_message.Send()


# Функция - 4: Спомогательная функция сортироваки массивов данных по ключу
def function_sort(number_value):
    def function_option(tuple_value):
        return tuple_value[number_value - 1]
    return function_option


#  Строковые константы
#  Строковая константа - 1: Параметры для подключения к удаленной БД (удалено из кода/скорректировано в коде)
status_db_string = """DRIVER=SQL Server;
                    SERVER=**************;
                    UID=*****************;
                    PWD=*****************;
                    DATABASE=*********"""

#  Строковая константа - 2: Параметры для подключения к удаленной БД (удалено из кода/скорректировано в коде)
plan_graph_db_string = """DRIVER=SQL Server;
                    SERVER=**************;
                    UID=*****************;
                    PWD=*****************;
                    DATABASE=*********"""

#  Строковая константа - 3: Список столбцов ПГ, для выгрузки или обновления данных
sql_string_pg = """region, nomer_pl, type_pl, status_pl, address, type_ams, po, freq_kls,
                date_inspect_fact, date_aop_plan, date_aop_fact, zakaz_pir_plan, zakaz_pir_fact,
                igi_plan, igi_fact, rns_plan, rns_fact, tu_energy_plan, tu_energy_fact, r1_plan,
                r1_fact, project_doc_plan, project_doc_fact, zakaz_smr_plan, zakaz_smr_fact,
                start_smr_plan, start_smr_fact, end_fundament_plan, end_fundament_fact,
                end_smr_ams_plan, end_smr_ams_fact, start_smr_hardware_plan, start_smr_hardware_fact,
                hardware_give_po_plan, hardware_give_po_fact, end_smr_hardware_plan, end_smr_hardware_fact,
                end_smr_plan, end_smr_fact, test_operation_plan, test_operation_fact, r2_plan, r2_fact,
                check_plan, check_fact, start_check_plan, start_check_fact, end_check_plan, end_check_fact,
                agreed_ks_2_3_plan, agreed_ks_2_3_fact, status_po, reason_block_factor,
                cypris, delay_aop, delay_zakaz_pir, delay_zakaz_smr, delay_start_smr, delay_end_smr,
                delay_test_operation, delay_start_check, delay_end_check, number_week_year, status_week"""

#  Строковая константа - 4: Список столбцов реестра задержек, для выгрузки или обновления данных
sql_string_delay = """po, region, week_year, count_bs,  delay_aop_previos, delay_aop_jan, delay_aop_feb,
            delay_aop_mar, delay_aop_apr, delay_aop_may, delay_aop_jun, delay_aop_jul, delay_aop_aug,
            delay_aop_sep, delay_aop_oct, delay_aop_nov, delay_aop_dec, delay_aop_sum, delay_pir_zakaz_previos,
            delay_pir_zakaz_jan, delay_pir_zakaz_feb, delay_pir_zakaz_mar, delay_pir_zakaz_apr, delay_pir_zakaz_may,
            delay_pir_zakaz_jun, delay_pir_zakaz_jul, delay_pir_zakaz_aug, delay_pir_zakaz_sep, delay_pir_zakaz_oct,
            delay_pir_zakaz_nov, delay_pir_zakaz_dec, delay_pir_zakaz_sum, delay_smp_zakaz_previos, delay_smp_zakaz_jan,
            delay_smp_zakaz_feb, delay_smp_zakaz_mar, delay_smp_zakaz_apr, delay_smp_zakaz_may, delay_smp_zakaz_jun,
            delay_smp_zakaz_jul, delay_smp_zakaz_aug, delay_smp_zakaz_sep, delay_smp_zakaz_oct, delay_smp_zakaz_nov,
            delay_smp_zakaz_dec, delay_smp_zakaz_sum, delay_smr_start_previos, delay_smr_start_jan, delay_smr_start_feb,
            delay_smr_start_mar, delay_smr_start_apr, delay_smr_start_may, delay_smr_start_jun, delay_smr_start_jul,
            delay_smr_start_aug, delay_smr_start_sep, delay_smr_start_oct, delay_smr_start_nov, delay_smr_start_dec,
            delay_smr_start_sum, delay_smr_end_previos, delay_smr_end_jan, delay_smr_end_feb, delay_smr_end_mar,
            delay_smr_end_apr, delay_smr_end_may, delay_smr_end_jun, delay_smr_end_jul, delay_smr_end_aug,
            delay_smr_end_sep, delay_smr_end_oct, delay_smr_end_nov, delay_smr_end_dec, delay_smr_end_sum, delay_to_fact_previos,
            delay_to_fact_jan, delay_to_fact_feb, delay_to_fact_mar, delay_to_fact_apr, delay_to_fact_may, delay_to_fact_jun,
            delay_to_fact_jul, delay_to_fact_aug, delay_to_fact_sep, delay_to_fact_oct, delay_to_fact_nov, delay_to_fact_dec,
            delay_to_fact_sum, delay_check_start_previos, delay_check_start_jan, delay_check_start_feb, delay_check_start_mar,
            delay_check_start_apr, delay_check_start_may, delay_check_start_jun, delay_check_start_jul, delay_check_start_aug,
            delay_check_start_sep, delay_check_start_oct, delay_check_start_nov, delay_check_start_dec, delay_check_start_sum,
            delay_check_end_previos, delay_check_end_jan, delay_check_end_feb, delay_check_end_mar, delay_check_end_apr,
            delay_check_end_may, delay_check_end_jun, delay_check_end_jul, delay_check_end_aug, delay_check_end_sep,
            delay_check_end_oct, delay_check_end_nov, delay_check_end_dec, delay_check_end_sum"""

#  Строковая константа - 5: Список столбцов общего реестра, для выгрузки или обновления данных
sql_string_reestr = """po, week, date_request, date_answer, type_attach,
                       date_repit_request, precent_error, count_error,
                       return_revision, need_aks"""

#  Строковая константа - 6: Параметры для выгрузки статусного отчета из БД (удалено из кода/скорректировано в коде)
status_select_value = """set dateformat dmy
                          select Регион, pl_name, con_type, pl, site_status, address, place_afu, org_smr, 
                          freq_kls, ACT_DATE_INSP, "Акт обследованияВыдан", "Заказ на ПИРПодписан", 
                          "Тех.условия на электроснабжение БС", Р1, "ПроектПередан", 
                          "Заказ на СМРПодписан", FACT_DATE_ST_CON,  "Готов к монтажу осн.оборуд.", 
                          "Основное оборудование получено", FACT_DATE_END_CON,  plan_date_bts, fact_date_bts, Р2, 
                          ACT_DATE_AC, "Акт РК (КС-11)Подписан", "Акт РК (КС-11)Подписан обеими сторонами", ЦУПРИС
                          from [*******].[**********]
                          where datepart(year, cast(plan_date_bts as date)) >= {tmp_year_start}
                          and datepart(year, cast(plan_date_bts as date)) <= {tmp_year_end}
                          and "Регион" != 'Москва'"""

#  Строковая константа - 7: Параметры для выгрузки статусного отчета из БД (удалено из кода/скорректировано в коде)
status_select_value_po = """set dateformat dmy
                          select Регион, pl_name, con_type, pl, site_status, address, place_afu, org_smr, 
                          freq_kls, ACT_DATE_INSP, "Акт обследованияВыдан", "Заказ на ПИРПодписан", 
                          "Тех.условия на электроснабжение БС", Р1, "ПроектПередан", 
                          "Заказ на СМРПодписан", FACT_DATE_ST_CON,  "Готов к монтажу осн.оборуд.", 
                          "Основное оборудование получено", FACT_DATE_END_CON,  plan_date_bts, fact_date_bts, Р2, 
                          ACT_DATE_AC, "Акт РК (КС-11)Подписан", "Акт РК (КС-11)Подписан обеими сторонами", ЦУПРИС
                          from [******].[**********]
                          where datepart(year, cast(plan_date_bts as date)) >= {tmp_year_start}
                          and datepart(year, cast(plan_date_bts as date)) <= {tmp_year_end}
                          and "Регион" != 'Москва' and org_smr = ('{po_name}')"""


#  Строковая константа - 8: Параметры для выгрузки операционного плана из БД (удалено из кода/скорректировано в коде)
freezing_select_value_pl = """select "№ Площадки", 'PL' AS pl_type, 'PL' AS freq_kls
                            from [*****].[*******]"""

#  Строковая константа - 9: Параметры для выгрузки операционного плана из БД (удалено из кода/скорректировано в коде)
freezing_select_value_cs = """SELECT pl_name , Площадка AS pl_type, freq_kls
                            FROM [*******].[********************]
                            WHERE Площадка IN ('CS', 'CS ID')"""

#  Строковая константа - 10: Шаблон письма для информирования руководства
html_value_1 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin: -10px; margin-left: -15px; margin-top: -20px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Коллеги, Добрый день!</p>
        <br>
        <p>Информируем, что по итогам работы с подрядными организация общим количеством <b>{count_total} шт.</b> по состоянию на <b>{week}</b>
        в части процесса получения план-графиков по объектам сети радиодоступа, получен следующий статус: </p>
        <p>1. Количество подрядных организаций, по которым получены корректные ПГ, составляет <b>{count_correct} шт.</b>;</p>
        <br>
        <p>2. Количество подрядных организаций, по которым в течение не более 2-х рассылок:
        <ul>
            <li>получены ПГ, не соответствующие требованию по формату - <b>{count_no_temp_1} шт.</b></li>
            <li>получены ПГ, с превышением порога ошибок - <b>{count_error_pg_1} шт.</b></p></li>
            <li>не получены ПГ - <b>{count_no_pg_1} шт.</b></p></li>
        </ul></p>
        <br>
        </body></html>
        """

#  Строковая константа - 11: Фрагментированный шаблон для информирования руководства (часть 1)
html_value_2_1 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px;}}
          p {{margin: 5px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>3. Список подрядных организаций, с которыми необходимо провести АКС по причинам:</p>
        </body></html>
        """

#  Строковая константа - 12: Фрагментированный шаблон для информирования руководства (часть 2)
html_value_2_2 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          table, th, td {{ border: 1px solid black;}}
          th, td {{ padding: 5px; }}
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px;}}
          p {{margin: 5px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li>получены ПГ не соответствующие требованию по формату в течении 2-х или более 
            рассылок по <b>{count_no_temp_2} шт.</b>: {table_1}</li><br>
        </ul>
        </body></html>
        """

#  Строковая константа - 13: Фрагментированный шаблон для информирования руководства (часть 3)
html_value_2_3 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          table, th, td {{ border: 1px solid black;}}
          th, td {{ padding: 5px; }}
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px;}}
          p {{margin: 5px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li>получены ПГ с превышением порога ошибок в течении 2-х или более 
            рассылок по <b>{count_error_pg_2} шт.</b>: {table_2}</li><br>
        </ul>
        </body></html>
        """

#  Строковая константа - 14: Фрагментированный шаблон для информирования руководства (часть 1)
html_value_2_4 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          table, th, td {{ border: 1px solid black;}}
          th, td {{ padding: 5px; }}
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px;}}
          p {{margin: 5px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li>не получены ПГ в течении 2-х или более рассылок по <b>{count_no_pg_2} шт.</b>: {table_3}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 15: Фрагментированный шаблон для информирования руководства (часть 1)
html_value_2_5 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px;}}
          p {{margin: 5px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <br>
        <p>Реестр ПО <b>(Reestr_cypris)</b> и реестр статусов просрочек <b>(Reestr_delay_cypris)</b> 
        по состоянию на <b>{week}</b> приведены во вложении:</p>
        <br>
        <p>C уважением,<br>
        </p>
        </body></html>
        """

#  Строковая константа - 16: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 1) (удалено из кода/скорректировано в коде)
html_value_3_1 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Добрый день!</p>
        <br>
        <p>********************************************************************************************.</p>
        <br>
        <p>На основании этого просим Вас, ответным письмом в течение 3-х дней <u>(в срок до {date_value_1})</u>, 
        предоставить план-график по выданным Вашей организации в работу объектам сети радиодоступа 
        (базовым станциям).</p>
        <br>
        </body></html>
        """

#  Строковая константа - 17: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 2)
html_value_3_2 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Обращаем внимание, что в соответствии с ранее предоставленными план - графиками от Вашей организацией, 
        а также имеющейся информацией о текущем статусе по строительству объектов сети радиодоступа, 
        выявлены просрочки*:</p>
        </body></html>
        """

#  Строковая константа - 18: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 3)
html_value_3_3 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Запуска БС в ТЭ в количестве {count_to} шт.:</b><br>
            {list_to}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 19: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 4)
html_value_3_4 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Выхода на СМР в количестве {count_smr} шт.:</b><br>
            {list_smr}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 20: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 5)
html_value_3_5 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Завершение СМР в количестве {count_smrg} шт.:</b><br>
            {list_smrg}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 21: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 6)
html_value_3_6 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Первичная приемка в количестве {count_check} шт.:</b><br>
            {list_check}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 22: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 7)
html_value_3_7 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Подписание КС-11 в количестве {count_ks} шт.:</b><br>
            {list_ks}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 23: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 8)
html_value_3_8 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Предоставления АОП в количестве {count_aop} шт.:</b><br>
            {list_aop}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 24: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 9)
html_value_3_9 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>При заполнении плана-графика просим Вас учесть изменения планируемых дат по этим объектам, указать причины 
        невозможности выполнения работ в срок в рамках договорных отношений, принятые/планируемые меры для устранения 
        отставания. </p>
        <br>
        </body></html>
        """

#  Строковая константа - 25: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (часть 10) (удалено из кода/скорректировано в коде)
html_value_3_10 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>При направлении ответа, важно:<br>
        1. Максимально точно и в полной мере, обращая внимание в первую очередь на подсвеченные ячейки, 
        заполнить в направляемом плане-графике фактические или планируемые даты (ознакомьтесь с инструкцией в 
        файле на листе «Прочитай!»)<br> 
        2. Указать обязательные адресаты: 
        <a href="mailto:******@****.****">*******@****.****</a> и 
        <a href="mailto:******@****.****">*******@****.****</a><br> 
        3. Поскольку обработка планов-графиков происходит в автоматическом режиме, НЕЛЬЗЯ: <br>
        - изменять название темы письма (при оправке в почтовом клиенте нажмите кнопку Ответить или Ответить всем) <br> 
        - изменять название и расширение (.xlsx) файла <br>
        - добавлять, удалять столбцы в файле <br>
        - архивировать файл </p>
        <br>
        <p>Если вы не отправили файл, отправили файл с нарушением этих требований, либо при заполнении допущено большое 
        кол-во ошибок, <u>{date_value_2}</u> Вы получите повторный запрос. </p>
        <br>
        <em>*Даты указанные во вложенном файле показывают на расчетный/нормативный/договорной срок, 
        к которому должен был быть выполнен соответствующий этап.</em></p> 
        <br>
        <p>C Уважением,<br>
        </p>
        </body></html>
        """

#  Строковая константа - 26: Форматированный шаблон письма для информирования региона (часть 1)
html_value_4_1 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Коллеги, Добрый день!</p>
        <br>
        <p>По итогу <b>{week_value} недели</b> информируем вас о получении от порядных организаций: <br>
        <b>{list_po}</b><br>
        план - графиков выполнения работ по объектам радиоподсистемы.<br>
        Просим вас ознакомиться с приложенными данными и при необходимости проработать с подрядными организациями 
        возможные корректировки плановых/фактических сроков выполнения работ.</p>
        <br>
        </body></html>
        """

#  Строковая константа - 27: Фрагментированный шаблон письма для информирования региона (часть 2)
html_value_4_2 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          table, th, td {{ border: 1px solid black;}}
          th, td {{ padding: 5px; }}
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Подрядные организации заявили о наличии блок-факторов, препятствующих строительству.<br>
        {table_4}</p>
        <br>
        </body></html>
        """

#  Строковая константа - 28: Фрагментированный шаблон письма для информирования региона (часть 3)
html_value_4_3 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          table, th, td {{ border: 1px solid black;}}
          th, td {{ padding: 5px; }}
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Подрядные организации заявили о завершении следующих этапов работ. Просим подтвердить факт выполнения работ 
        и отразить фактические даты.<br>
        {table_1}</p>
        <br>
        </body></html>
        """

#  Строковая константа - 29: Фрагментированный шаблон письма для информирования региона (часть 4)
html_value_4_4 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          table, th, td {{ border: 1px solid black;}}
          th, td {{ padding: 5px; }}
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Подрядные организации заявили о наличии в работе площадок, которые не отражены в план-графиках при запросе. 
        Просим подтвердить и выполнить распределение, а также отразить информацию о фактических датах 
        выполнения работ.<br>
        {table_2}</p>
        <br>
        </body></html>
        """

#  Строковая константа - 30: Фрагментированный шаблон письма для информирования региона (часть 5)
html_value_4_5 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          table, th, td {{ border: 1px solid black;}}
          th, td {{ padding: 5px; }}
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Подрядные организации заявили о наличии площадок, которые не передавались им в работу. 
        Просим проверить и выполнить перераспределение.<br>
        {table_3}</p>
        <br>
        </body></html>
        """

#  Строковая константа - 31: Фрагментированный шаблон письма для информирования региона (часть 6)
html_value_4_6 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          table, th, td {{ border: 1px solid black; border-collapse: collapse;}}
          th, td {{ padding: 5px; }}
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>C уважением,<br>
        </p>
        </body></html>
        """

#  Строковая константа - 32: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 1) (удалено из кода/скорректировано в коде)
html_value_5_1 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Добрый день!</p>
        <br>
        <p>*********************************************************************************************<br>
        <br>
        <b><u>{date_value_1}  в адрес  Вашей подрядной организации был направлен соответствующий запрос, но, 
        к сожалению, план-график <font color='red'>{status}.</font></u></b></p>
        <br>
        </body></html>
        """

#  Строковая константа - 33: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 2)
html_value_5_2 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>Обращаем внимание, что в соответствии с ранее предоставленными план - графиками от Вашей организацией, 
        а также информацией о текущем статусе по строительству объектов сети радиодоступа, 
        выявлены просрочки*:</p>
        </body></html>
        """

#  Строковая константа - 34: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 3)
html_value_5_3 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Запуска БС в ТЭ в количестве {count_to} шт.:</b><br>
            {list_to}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 35: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 4)
html_value_5_4 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Выхода на СМР в количестве {count_smr} шт.:</b><br>
            {list_smr}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 36: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 5)
html_value_5_5 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Завершение СМР в количестве {count_smrg} шт.:</b><br>
            {list_smrg}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 37: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 6)
html_value_5_6 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Первичная приемка в количестве {count_check} шт.:</b><br>
            {list_check}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 38: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 7)
html_value_5_7 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Подписание КС-11 в количестве {count_ks} шт.:</b><br>
            {list_ks}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 39: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 8)
html_value_5_8 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <ul>
            <li><b>Предоставления АОП в количестве {count_aop} шт.:</b><br>
            {list_aop}</li>
        </ul>
        </body></html>
        """

#  Строковая константа - 40: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 9)
html_value_5_9 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>При заполнении плана-графика просим Вас учесть изменения планируемых дат по этим объектам, указать причины 
        невозможности выполнения работ в срок в рамках договорных отношений, принятые/планируемые меры для устранения 
        отставания. </p>
        </body></html>
        """

#  Строковая константа - 41: Фрагментированный шаблон письма для запроса ПГ в адрес ПО (повторный запрос - часть 10) (удалено из кода/скорректировано в коде)
html_value_5_10 = """
        <html>
        <head>
        <meta charset="utf-8">
        <style>
          ul {{margin-top: -10px;}}
          li {{margin-left: -15px; margin-top: 5px;}}
          p {{margin: 2px;}}
          br {{line-height: 80%;}}
        </style>
        </head>
        <html><body>
        <p>При направлении ответа, важно:<br>
        1. Максимально точно и в полной мере, обращая внимание в первую очередь на подсвеченные ячейки, 
        заполнить в направляемом плане-графике фактические или планируемые даты (ознакомьтесь с инструкцией в 
        файле на листе «Прочитай!»)<br> 
        2. Указать обязательные адресаты: 
        <a href="mailto:******@****.****">*******@****.****</a> и 
        <a href="mailto:******@****.****">*******@****.****</a><br> 
        3. Поскольку обработка планов-графиков происходит в автоматическом режиме, НЕЛЬЗЯ: <br>
        - изменять название темы письма (при оправке в почтовом клиенте нажмите кнопку Ответить или Ответить всем) <br> 
        - изменять название и расширение (.xlsx) файла <br>
        - добавлять, удалять столбцы в файле <br>
        - архивировать файл </p>
        <br>
        <p>В связи с этим, повторно просим Вас, ответным письмом в течение 2-х дней, <b><u>до {date_value_2}</u></b>, предоставить 
        в план-график по выданным Вашей организации в работу объектам сети радиодоступа (базовым станциям). 
        Письма, полученные по истечении 2-х дневного срока, автоматически не будут обработаны и приняты к учету. 
        Если Вам не ясна причина повторного запроса, нужно написать обращение на e-mail: 
        <a href="mailto:******@****.****">*******@****.****</a> и 
        <a href="mailto:******@****.****">*******@****.****</a><br>  
        для уточнения и консультации. </p>
        <br>
        <em>*Даты указанные во вложенном файле показывают на расчетный/нормативный/договорной срок, 
        к которому должен был быть выполнен соответствующий этап.</em></p> 
        <br>
        <p>C уважением,<br>
        </p>
        </body></html>
        """

#  Строковая константа - 42: Сообщение об ошибке при проверке валидности данных
error_value_date = 'Пожалуйста, укажите плановую/фактическую дату выполнения работ по данному этапу. Для плановых не принимается прошлый периода. Для фактических не принимаются будущий периода. Для комментариев/причины остановки строительства есть отдельные столбцы справа.'

#  Строковая константа - 43: Сообщение об ошибке при проверке валидности данных
error_value_list = 'Пожалуйста, выберите значение из выпадающего списка'

#  Строковая константа - 44: Запрос SQL на выгрузку ПГ за последнюю корректную неделю
pandas_sql_value_3 = """select * from database_pg
                        where nomer_pl = ('{pl}')
                        and po = ('{po}')
                        and number_week_year = ('{week_year}');"""

#  Кортежные константы
#  Кортежная константа - 1: Кортеж индексов ПГ из БД
iteration_tuple = (7, 9, 11, 13, 15, 17, 19, 21,
                   23, 25, 27, 29, 31, 33, 35,
                   37, 39, 41, 43, 45, 47, 49)

#  Кортежная константа - 2: Кортеж для заполнения общего реестра
template_value = ('ответ не получен', 'не по шаблону')

#  Кортежная константа - 3: Кортеж из списка ПО требующих исключение (удалено из кода/скорректировано в коде)
exept_list_po = ('None', '')

#  Кортежная константа - 4: Кортеж основных направлений строительства
new_rsr_list = ('Строительство новой площадки (ПК,ПЗ)',
                'Строительство нового диапазона (ДС)')

#  Кортежная константа - 5: Кортеж для формирования запроса загрузки ПГ в БД
sql_tuple_1 = ('region', 'nomer_pl', 'type_pl', 'status_pl', 'address',
               'type_ams', 'po', 'freq_kls', 'date_inspect_fact',
               'date_aop_plan', 'date_aop_fact', 'zakaz_pir_plan',
               'zakaz_pir_fact', 'igi_plan', 'igi_fact', 'rns_plan',
               'rns_fact', 'tu_energy_plan', 'tu_energy_fact', 'r1_plan',
               'r1_fact', 'project_doc_plan', 'project_doc_fact',
               'zakaz_smr_plan', 'zakaz_smr_fact', 'start_smr_plan',
               'start_smr_fact', 'end_fundament_plan', 'end_fundament_fact',
               'end_smr_ams_plan', 'end_smr_ams_fact', 'start_smr_hardware_plan',
               'start_smr_hardware_fact', 'hardware_give_po_plan',
               'hardware_give_po_fact', 'end_smr_hardware_plan',
               'end_smr_hardware_fact', 'end_smr_plan', 'end_smr_fact',
               'test_operation_plan', 'test_operation_fact', 'r2_plan',
               'r2_fact', 'check_plan', 'check_fact', 'start_check_plan', 'start_check_fact',
               'end_check_plan', 'end_check_fact', 'agreed_ks_2_3_plan',
               'agreed_ks_2_3_fact', 'status_po', 'reason_block_factor', 'cypris')

#  Кортежная константа - 6: Кортеж для формирования запроса загрузки ПГ в БД
sql_tuple_2 = ('delay_aop', 'delay_zakaz_pir', 'delay_zakaz_smr', 'delay_start_smr', 'delay_end_smr',
               'delay_test_operation', 'delay_start_check', 'delay_end_check', 'number_week_year')

#  Кортежная константа - 7 Вспомогательный кортеж для значения словаря
sql_header_reestr = ('po', 'week', 'date_request', 'date_answer', 'type_attach',
                     'date_repit_request', 'precent_error', 'count_error',
                     'return_revision', 'need_aks')

#  Кортежная константа - 8 Вспомогательный кортеж для значения словаря
sql_header_delay = ('po', 'region', 'count_bs', 'aop', 'pir_order',
                    'smr_order', 'start_smr', 'end_smr', 'test_operation',
                    'start_check', 'end_check', 'week_year')

#  Кортежная константа - 9: Кортеж списка БД размороженных объектов (удалено из кода/скорректировано в коде)
freezing_list = ''

#  Кортежная константа - 10: Кортеж индексов сводного реестра задержек в формате .csv
send_index_list = (3, 17, 31, 45, 59, 73, 87, 101, 115)

#  Словарные константы
#  Словарная константа - 1: Словарь индексов ПГ
iteration_dict = {0: 0, 1: 1, 2: 3, 3: 4, 4: 5, 5: 6, 7: 8, 8: 9, 10: 10,
                  12: 11, 18: 12, 20: 13, 22: 14, 24: 15, 26: 16, 32: 17,
                  34: 18, 38: 19, 40: 21, 42: 22, 46: 23, 48: (24, 25), 53: 26}

#  Словарная константа - 2: Словарь индексов ПГ
iteration_dict_upd = {10: 10, 12: 11, 18: 12, 20: 13, 22: 14, 24: 15, 26: 16,
                      32: 17, 34: 18, 38: 19, 40: 21, 42: 22, 46: 23, 48: 24}

#  Словарная константа - 3: Словарь для формирования реестра задержек
check_dict = {0: (8, 10, 'Выдачи АОП', 10), 1: (10, 12, 'Выдачи заказов на ПИР', 20),
              2: (10, 24, 'Выдачи заказов на СМР', 20), 3: (12, 26, 'Факт начало СМР', (30, 30, 40)),
              4: (24, 38, 'Факт окончание СМР', (15, 30, 45)), 5: (38, 40, 'Факт ТЭ', 7),
              6: (38, 46, 'Выезд на приемку', 15), 7: (46, 48, 'Подписание КС-11', 30)}

#  Словарная константа - 4: Словарь локальныйх БД (удалено из кода/скорректировано в коде)
sql_table_backup = {}

#  Словарная константа - 5: Словарь локальныйх БД (удалено из кода/скорректировано в коде)
sql_table_header = {}

#  Словарная константа - 6: Словарь формул расчета ошибок
excel_dict = {
    'BC7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",J7,1),0),IFERROR(SEARCH("треб",K7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((K7*1),5)=5,IFERROR((J7*1),5)=5),1,0)+IF(AND(OR(K7="",IFERROR((K7*1),5)=5),IFERROR((J7*1),5)<$CA$1),1,0)+IF(IFERROR((K7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BD7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",L7,1),0),IFERROR(SEARCH("треб",M7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((M7*1),5)=5,IFERROR((L7*1),5)=5),1,0)+IF(AND(OR(M7="",IFERROR((M7*1),5)=5),IFERROR((L7*1),5)<$CA$1),1,0)+IF(IFERROR((M7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BE7': '=IF(OR($C7="CS",$C7="RT",$C7="ID"),0,IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",N7,1),0),IFERROR(SEARCH("треб",O7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((O7*1),5)=5,IFERROR((N7*1),5)=5),1,0)+IF(AND(OR(O7="",IFERROR((O7*1),5)=5),IFERROR((N7*1),5)<$CA$1),1,0)+IF(IFERROR((O7*1),5)>$CA$1,1,0))>=1,1,0))),0),0))',
    'BF7': '=IF($C7="GF",0,IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",P7,1),0),IFERROR(SEARCH("треб",Q7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((Q7*1),5)=5,IFERROR((P7*1),5)=5),1,0)+IF(AND(OR(Q7="",IFERROR((Q7*1),5)=5),IFERROR((P7*1),5)<$CA$1),1,0)+IF(IFERROR((Q7*1),5)>$CA$1,1,0))>=1,1,0))),0),0))',
    'BG7': '=IF($C7="CS",0,IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",R7,1),0),IFERROR(SEARCH("треб",S7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((S7*1),5)=5,IFERROR((R7*1),5)=5),1,0)+IF(AND(OR(S7="",IFERROR((S7*1),5)=5),IFERROR((R7*1),5)<$CA$1),1,0)+IF(IFERROR((S7*1),5)>$CA$1,1,0))>=1,1,0))),0),0))',
    'BH7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",T7,1),0),IFERROR(SEARCH("треб",U7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((U7*1),5)=5,IFERROR((T7*1),5)=5),1,0)+IF(AND(OR(U7="",IFERROR((U7*1),5)=5),IFERROR((T7*1),5)<$CA$1),1,0)+IF(IFERROR((U7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BI7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",V7,1),0),IFERROR(SEARCH("треб",W7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((W7*1),5)=5,IFERROR((V7*1),5)=5),1,0)+IF(AND(OR(W7="",IFERROR((W7*1),5)=5),IFERROR((V7*1),5)<$CA$1),1,0)+IF(IFERROR((W7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BJ7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",X7,1),0),IFERROR(SEARCH("треб",Y7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((Y7*1),5)=5,IFERROR((X7*1),5)=5),1,0)+IF(AND(OR(Y7="",IFERROR((Y7*1),5)=5),IFERROR((X7*1),5)<$CA$1),1,0)+IF(IFERROR((Y7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BK7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",Z7,1),0),IFERROR(SEARCH("треб",AA7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AA7*1),5)=5,IFERROR((Z7*1),5)=5),1,0)+IF(AND(OR(AA7="",IFERROR((AA7*1),5)=5),IFERROR((Z7*1),5)<$CA$1),1,0)+IF(IFERROR((AA7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BL7': '=IF(OR($C7="CS",$C7="RT",$C7="ID"),0,IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",AB7,1),0),IFERROR(SEARCH("треб",AC7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AC7*1),5)=5,IFERROR((AB7*1),5)=5),1,0)+IF(AND(OR(AC7="",IFERROR((AC7*1),5)=5),IFERROR((AB7*1),5)<$CA$1),1,0)+IF(IFERROR((AC7*1),5)>$CA$1,1,0))>=1,1,0))),0),0))',
    'BM7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",AD7,1),0),IFERROR(SEARCH("треб",AE7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AE7*1),5)=5,IFERROR((AD7*1),5)=5),1,0)+IF(AND(OR(AE7="",IFERROR((AE7*1),5)=5),IFERROR((AD7*1),5)<$CA$1),1,0)+IF(IFERROR((AE7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BN7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",AF7,1),0),IFERROR(SEARCH("треб",AG7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AG7*1),5)=5,IFERROR((AF7*1),5)=5),1,0)+IF(AND(OR(AG7="",IFERROR((AG7*1),5)=5),IFERROR((AF7*1),5)<$CA$1),1,0)+IF(IFERROR((AG7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BO7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",AH7,1),0),IFERROR(SEARCH("треб",AI7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AI7*1),5)=5,IFERROR((AH7*1),5)=5),1,0)+IF(AND(OR(AI7="",IFERROR((AI7*1),5)=5),IFERROR((AH7*1),5)<$CA$1),1,0)+IF(IFERROR((AI7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BP7': '=IF($BA7="",IF(OR($AM7="",IFERROR(($AM7*1),"")=""),IF(SUM(IFERROR(SEARCH("треб",AJ7,1),0),IFERROR(SEARCH("треб",AK7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AK7*1),5)=5,IFERROR((AJ7*1),5)=5),1,0)+IF(AND(OR(AK7="",IFERROR((AK7*1),5)=5),IFERROR((AJ7*1),5)<$CA$1),1,0)+IF(IFERROR((AK7*1),5)>$CA$1,1,0))>=1,1,0))),0),0)',
    'BQ7': '=IF($BA7="",IF(SUM(IFERROR(SEARCH("треб",AL7,1),0),IFERROR(SEARCH("треб",AM7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AM7*1),5)=5,IFERROR((AL7*1),5)=5),1,0)+IF(AND(OR(AM7="",IFERROR((AM7*1),5)=5),IFERROR((AL7*1),5)<$CA$1),1,0)+IF(IFERROR((AM7*1),5)>$CA$1,1,0))>=1,1,0))),0)',
    'BR7': '=IF($BA7="",IF(SUM(IFERROR(SEARCH("треб",AN7,1),0),IFERROR(SEARCH("треб",AO7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AO7*1),5)=5,IFERROR((AN7*1),5)=5),1,0)+IF(AND(OR(AO7="",IFERROR((AO7*1),5)=5),IFERROR((AN7*1),5)<$CA$1),1,0)+IF(IFERROR((AO7*1),5)>$CA$1,1,0))>=1,1,0))),0)',
    'BS7': '=IF($BA7="",IF(SUM(IFERROR(SEARCH("треб",AP7,1),0),IFERROR(SEARCH("треб",AQ7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AQ7*1),5)=5,IFERROR((AP7*1),5)=5),1,0)+IF(AND(OR(AQ7="",IFERROR((AQ7*1),5)=5),IFERROR((AP7*1),5)<$CA$1),1,0)+IF(IFERROR((AQ7*1),5)>$CA$1,1,0))>=1,1,0))),0)',
    'BT7': '',
    'BU7': '=IF($BA7="",IF(SUM(IFERROR(SEARCH("треб",AT7,1),0),IFERROR(SEARCH("треб",AU7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AU7*1),5)=5,IFERROR((AT7*1),5)=5),1,0)+IF(AND(OR(AU7="",IFERROR((AU7*1),5)=5),IFERROR((AT7*1),5)<$CA$1),1,0)+IF(IFERROR((AU7*1),5)>$CA$1,1,0))>=1,1,0))),0)',
    'BV7': '=IF($BA7="",IF(SUM(IFERROR(SEARCH("треб",AV7,1),0),IFERROR(SEARCH("треб",AW7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AW7*1),5)=5,IFERROR((AV7*1),5)=5),1,0)+IF(AND(OR(AW7="",IFERROR((AW7*1),5)=5),IFERROR((AV7*1),5)<$CA$1),1,0)+IF(IFERROR((AW7*1),5)>$CA$1,1,0))>=1,1,0))),0)',
    'BW7': '=IF($BA7="",IF(SUM(IFERROR(SEARCH("треб",AX7,1),0),IFERROR(SEARCH("треб",AY7,1),0))>=1,0,(IF(SUM(IF(AND(IFERROR((AY7*1),5)=5,IFERROR((AX7*1),5)=5),1,0)+IF(AND(OR(AY7="",IFERROR((AY7*1),5)=5),IFERROR((AX7*1),5)<$CA$1),1,0)+IF(IFERROR((AY7*1),5)>$CA$1,1,0))>=1,1,0))),0)',
    'BX7': '=IF((SUM(BC7:BW7)/COUNTA(BC7:BW7))<=$BY$1,"Без ошибок","Дозаполнить")',
    'BY7': '=COUNTIF(BC7:BW7,1)/COUNTA(BC7:BW7)'
}

#  Словарная константа - 7: Словарь этапов ПГ
dict_stage_pg = {
    8: 'Дата обследования',
    10: 'Акт обследования',
    12: 'Заказ на ПИР',
    20: 'Р1',
    22: 'Проект передан',
    24: 'Заказ на СМР',
    26: 'Начало СМР',
    38: 'Окончание СМР',
    40: 'Факт ТЭ',
    42: 'Р2',
    46: 'Выезд на приёмку',
    48: 'Подписание КС-11'
}

#  Словарная константа - 8: Словарь стурктуры (удалено из кода/скорректировано в коде)
dict_cypris = {}
