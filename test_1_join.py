from apscheduler.schedulers.blocking import BlockingScheduler
from outlook_templates import art_print_start
from test_1_multiply_send import main_send_1
from test_2_multiply_send import main_send_2
from test_3_multiply_send import main_send_3
from test_1_multiply_get import main_get_1
from test_1_xl import main_xl_1
from test_2_xl import main_xl_2
from test_3_xl import main_xl_3
from test_1_db import main_db
from test_1_bg import main_bg


# Скрипт по запуску всех скриптов outlook_scripts в соответствии с заданным расписанием
# Разработчик: Ilya Trulkanovich
# Статус: Тестирование

# ----------------------------
# Этапы (оглавление):
# 1 этап: Функции/программы по автоматизированной работе с ПО (Общее кол-во: 4 функции)
# 2 этап: Запуск основной функции/программы

# 2 этап: Функции
# Функция - 1: Функция понедельника, рассылка итогов в адрес руководства ЦУПРИС и сотрудников ОРС (ТД + НОРС)
def function_monday():
    main_send_2()
    main_send_3()


# Функция - 2: Функция четверга, первичный анализ полученных ПГ от ПО, формирование повторной рассылки/запросов ПГ
def function_thursday():
    main_get_1()
    main_xl_2()


# Функция - 3: Функция пятницы, формирования, архивирование, отправка ПГ в адрес ПО
def function_friday():
    main_xl_1()
    main_xl_3()
    main_send_1()


# Функция - 4: Функция субботы, повторный анализ полученных ПГ от ПО, запись ПГ и реестров в БД (ДКРИС + локальная)
def function_saturday():
    main_get_1()
    main_xl_2()
    main_bg()
    main_db()


# 2 этап: Запуск основной функции/программы
def main():
    art_print_start(name_project='Start project')
    scheduler_value = BlockingScheduler()
    scheduler_value.add_job(function_friday, 'interval', weeks=2, start_date='2022-11-18 14:00:00',
                            timezone='Europe/Moscow')
    scheduler_value.add_job(function_thursday, 'interval', weeks=2, start_date='2022-11-24 05:00:00',
                            timezone='Europe/Moscow')
    scheduler_value.add_job(function_saturday, 'interval', weeks=2, start_date='2022-11-26 05:00:00',
                            timezone='Europe/Moscow')
    scheduler_value.add_job(function_monday, 'interval', weeks=2, start_date='2022-11-28 05:00:00',
                            timezone='Europe/Moscow')
    scheduler_value.start()


if __name__ == '__main__':
    main()
