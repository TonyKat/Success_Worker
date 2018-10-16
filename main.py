import os
import re
import sys
import xlrd
from datetime import datetime


class Worker:
    """
    "Успешные" дни (быстродействие) = План. - Факт.
    n_lead - количество руководящих должностей (быть руководителем проекта)
    n_success_delivery - количество успешных сдач проектов в руководящей должности
    n_success_lead_days - количество успешных дней в руководящей должности
    n_projects - количество участий в проектах (за вычетом участий в проектах в руков-й должности)
    n_success_days - количество успешных дней всего (за вычетом успешных дней в руков-й должности)
    """

    def __init__(self,
                 n_lead=None,
                 n_success_delivery=None,
                 n_success_lead_days=None,
                 n_projects=None,
                 n_success_days=None):
        self.n_lead = []
        self.n_success_delivery = 0
        self.n_success_lead_days = 0
        self.n_projects = []
        self.n_success_days = 0


def input_data(regex, text):
    matches = re.findall(regex, input(text))
    while not matches:
        try:
            print('Некорректные данные!')
            matches = re.findall(regex, input(text))
        except:
            print('Некорректные данные!')
    return matches[0]


def get_xlsx_path(loc_xlsx_dir):
    loc_xlsx_path = []
    for file in os.listdir(loc_xlsx_dir):
        if file.endswith('.xlsx'):
            loc_xlsx_path.append(os.path.join(loc_xlsx_dir, file))
    return loc_xlsx_path


def input_path():
    while True:
        absolute_path = input('Введите абсолютный путь к директории с excel-файлами (для отмены ввода нажмите "0"): ')
        if absolute_path == '0':
            break
        try:
            if os.path.exists(absolute_path) is True:
                return get_xlsx_path(absolute_path)
            else:
                print('Не найдена директория!')
                continue
        except:
            print('Ошибка поиска файла!')

    if absolute_path == '0':
        sys.exit()
    else:
        return get_xlsx_path(absolute_path)


def get_names(loc_xlsx_path):
    loc_d_worker = dict()
    for file in loc_xlsx_path:
        data = xlrd.open_workbook(filename=file)
        table = data.sheets()[0]

        # пробежаться по 0 строке и добавить в словарь все имена
        row_slice = table.row_slice(0, 4)
        for i in range(len(row_slice)):
            value = row_slice[i].value[:-6]
            if loc_d_worker.get(value) is None:
                loc_d_worker[value] = Worker()

        # пробежаться по 1 столбцу и добавить в словарь все имена
        col_slice = table.col_slice(1, 1)
        for i in range(len(col_slice)):
            value = col_slice[i].value
            if loc_d_worker.get(value) is None:
                loc_d_worker[value] = Worker()
    return loc_d_worker


def get_info(loc_xlsx_path, worker):
    for file in loc_xlsx_path:
        data = xlrd.open_workbook(filename=file)
        table = data.sheets()[0]

        for row in range(1, table.nrows):
            for col in range(0, table.ncols, 2):
                value = table.cell(row, col).value

                if col == 0:
                    # подсчет проектов в руководящей должности
                    try:
                        if value not in worker.get(table.cell(row, 1).value).n_lead:
                            worker.get(table.cell(row, 1).value).n_lead.append(value)
                    except:
                        print('Ошибка при добавлении в n_lead.')
                elif col == 2:
                    # подсчет успешных сдач проектов в руководящей должности
                    try:
                        date_plan = datetime(*xlrd.xldate_as_tuple(float(value), data.datemode))
                        date_fact = datetime(*xlrd.xldate_as_tuple(float(table.cell(row, col + 1).value),
                                                                   data.datemode))
                        days = (date_plan - date_fact).days
                        if days > 0 or days == 0:
                            worker[table.cell(row, 1).value].n_success_delivery += 1
                    except:
                        print('Ошибка при работе с датами.')
                else:
                    # подсчет успешных дней (План. - Факт.) участников проекта
                    try:
                        # если человек не участвовал в проекте - пропустить
                        if (value == '' or value is None or value == 0) \
                                and (table.cell(row, col + 1).value == '' or
                                     table.cell(row, col + 1).value is None or
                                     table.cell(row, col + 1).value == 0):
                            continue
                        if value == '' or value is None:
                            value = 0
                        if table.cell(row, col + 1).value == '' or table.cell(row, col + 1).value is None:
                            days = int(value)
                        else:
                            days = int(value) - int(table.cell(row, col + 1).value)

                        # если имя руководителя проекта не совпадает с именем участника проекта: успешные дни += days
                        if table.cell(0, col).value[:-6] != table.cell(row, 1).value:
                            worker[table.cell(0, col).value[:-6]].n_success_days += days
                            # если названия проекта нет в списке проектов участника проекта: проекты += 1
                            if table.cell(row, 0).value not in worker[table.cell(0, col).value[:-6]].n_projects:
                                worker[table.cell(0, col).value[:-6]].n_projects.append(table.cell(row, 0).value)
                        else:  # имя руководителя проекта совпадает с именем участника проекта
                            # успешные дни как руководителя проекта += days
                            worker[table.cell(0, col).value[:-6]].n_success_lead_days += days
                            # если названия проекта нет в списке лидерских проектов: лидерские проекты += 1
                            if table.cell(row, 0).value not in worker[table.cell(0, col).value[:-6]].n_lead:
                                worker[table.cell(0, col).value[:-6]].n_lead.append(table.cell(row, 0).value)
                    except:
                        print('Ошибка при подсчете успешных дней.')

    return worker


def print_list(w):
    print()
    for i in range(len(w)):
        print('{}. '
              'Name = {}'
              '\nn_lead = {}'
              '\nn_success_delivery = {}'
              '\nn_success_lead_days = {}'
              '\nn_projects = {}'
              '\nn_success_days = {}\n'
              .format(i + 1,
                      w[i][0],
                      len(w[i][1].n_lead),
                      w[i][1].n_success_delivery,
                      w[i][1].n_success_lead_days,
                      len(w[i][1].n_projects),
                      w[i][1].n_success_days))
    return


def sorted_workers(workers):
    try:
        data = input_data(r'^[0123456]{1}$',
                          '\nВыберите критерий вывода:'
                          '\n1. Составной критерий (все критерии сразу), нажмите "1".'
                          '\n2. По количеству руководящих должностей, нажмите "2".'
                          '\n3. По количеству успешных сдач проектов в руководящей должности, нажмите "3".'
                          '\n4. По количеству успешных дней в руководящей должности, нажмите "4".'
                          '\n5. По количеству участий в проектах (без лидерских проектов), нажмите "5".'
                          '\n6. По количеству успешных дней в проектах (без учета успешных лидерских дней), '
                          'нажмите "6". '
                          '\n7. Выйти, нажмите "0".'
                          '\nВведите: ')
        if data == '1':
            workers_list = sorted(workers.items(), key=lambda item: (len(item[1].n_lead),
                                                                     item[1].n_success_delivery,
                                                                     item[1].n_success_lead_days,
                                                                     len(item[1].n_projects),
                                                                     item[1].n_success_days), reverse=True)
        elif data == '2':
            workers_list = sorted(workers.items(), key=lambda item: len(item[1].n_lead), reverse=True)
        elif data == '3':
            workers_list = sorted(workers.items(), key=lambda item: item[1].n_success_delivery, reverse=True)
        elif data == '4':
            workers_list = sorted(workers.items(), key=lambda item: item[1].n_success_lead_days, reverse=True)
        elif data == '5':
            workers_list = sorted(workers.items(), key=lambda item: len(item[1].n_projects), reverse=True)
        elif data == '6':
            workers_list = sorted(workers.items(), key=lambda item: item[1].n_success_days, reverse=True)
        else:
            print('\nПереход в главное меню...')
            return '0'

        print_list(workers_list)
        return data
    except:
        print('\nПереход в главное меню...')
    return


def main_menu():
    try:
        while True:
            xlsx_path = input_path()
            if len(xlsx_path) == 0:
                print('Не найдены excel-файлы!\n')
                continue
            d_worker = get_names(xlsx_path)
            d_worker = get_info(xlsx_path, d_worker)
            while True:
                num = sorted_workers(d_worker)
                if num == '0':
                    break
            main_num = input_data(r'^[01]{1}$', '\nВвести другой путь к excel-файлам?'
                                                '\n"Да", нажмите "1".'
                                                '\n"Нет", нажмите "0".'
                                                '\nВведите: ')
            if main_num == '0':
                print('\nДо свидания...')
                break
    except:
        print('\nДо свидания...')

    return


if __name__ == '__main__':
    main_menu()
