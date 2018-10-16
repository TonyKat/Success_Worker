from datetime import datetime
import os
import re
import xlrd


class Worker:
    """
    "Успешные" дни = План. - Факт.
    name - ФИО сотрудника
    n_lead - число руководящих должностей (быть руководителем проекта)
    n_success_delivery - число успешных сдач проектов в руководящей должности
    n_success_lead_days - число успешных дней в руководящей должности
    n_projects - число участий в проектах (за вычетом участий в проектах в руков-й должности)
    n_success_days - число успешных дней всего (за вычетом успешных дней в руков-й должности)
    """

    def __init__(self,
                 name,
                 n_lead=None,
                 n_success_delivery=None,
                 n_success_lead_days=None,
                 n_projects=None,
                 n_success_days=None):
        self.name = name
        self.n_lead = []
        self.n_success_delivery = 0
        self.n_success_lead_days = 0
        self.n_projects = []
        self.n_success_days = 0


def get_xlsx_path(loc_xlsx_dir):
    loc_xlsx_path = []
    for file in os.listdir(loc_xlsx_dir):
        if file.endswith('.xlsx'):
            loc_xlsx_path.append(os.path.join(loc_xlsx_dir, file))
    return loc_xlsx_path


def get_names(loc_xlsx_path):
    loc_d_worker = dict()
    for file in loc_xlsx_path:
        data = xlrd.open_workbook(filename=file)
        table = data.sheets()[0]

        # row_slice(rowx, start, end) - срез ячеек(номер строки, начало, конец)
        row_slice = table.row_slice(0, 4)
        for i in range(len(row_slice)):
            value = row_slice[i].value[:-6]
            if loc_d_worker.get(value) is None:
                loc_d_worker[value] = Worker(name=value)

        col_slice = table.col_slice(1, 1)
        for i in range(len(col_slice)):
            value = col_slice[i].value
            if loc_d_worker.get(value) is None:
                loc_d_worker[value] = Worker(name=value)
    return loc_d_worker


def get_info(loc_xlsx_path, worker):
    for file in loc_xlsx_path:
        data = xlrd.open_workbook(filename=file)
        table = data.sheets()[0]

        for row in range(1, table.nrows):
            for col in range(0, table.ncols, 2):
                value = table.cell(row, col).value

                if col == 0:
                    # подсчет проектов в руков-й должности
                    try:
                        if value not in worker.get(table.cell(row, 1).value).n_lead:
                            worker.get(table.cell(row, 1).value).n_lead.append(value)
                    except:
                        print('Ошибка при добавлении в n_lead.')
                elif col == 2:
                    # подсчет успешных сдач проектов в руков-й должности
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
                    # подсчет успешных дней
                    try:
                        # если человек не участвовал в проекте - пропустить
                        if (value == '' or value is None) \
                                and (table.cell(row, col + 1).value == '' or table.cell(row, col + 1).value is None):
                            continue
                        if value == '' or value is None:
                            value = 0
                        if table.cell(row, col + 1).value == '' or table.cell(row, col + 1).value is None:
                            days = int(value)
                        else:
                            days = int(value) - int(table.cell(row, col + 1).value)

                        # если имя руководителя проекта не совпадает с участником проекта: успешные дни += days
                        if table.cell(0, col).value[:-6] != table.cell(row, 1).value:
                            worker[table.cell(0, col).value[:-6]].n_success_days += days
                            # если названия проекта нет в списке проекта участника проекта: проекты += 1
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


def input_data(regex, text):
    matches = re.findall(regex, input(text))
    while not matches:
        try:
            print('Некорректные данные!')
            matches = re.findall(regex, input(text))
        except:
            print('Некорректные данные!')
    return matches[0]


def print_sorted(workers):
    try:
        data = input_data(r'^[012345]{1}$',
                          '\nВыберите вывод:'
                          '\n1) По числу руководящих должностей, нажмите "1".'
                          '\n2) По числу успешных сдач проектов в руководящей должности, нажмите "2".'
                          '\n3) По числу успешных дней в руководящей должности, нажмите "3".'
                          '\n4) По числу участий в проектах (без лидерских проектов), нажмите "4".'
                          '\n5) По числу успешных дней в проектах (без учета успешных лидерских дней), нажмите "5".'
                          '\n6) Выйти, нажмите "0".'
                          '\nВведите: ')
        if data == '1':
            # n_lead = ['Проект1', 'Проект2']
            workers_list = sorted(workers.items(), key=lambda item: len(item[1].n_lead), reverse=True)
        elif data == '2':
            workers_list = sorted(workers.items(), key=lambda item: item[1].n_success_delivery, reverse=True)
        elif data == '3':
            workers_list = sorted(workers.items(), key=lambda item: item[1].n_success_lead_days, reverse=True)
        elif data == '4':
            workers_list = sorted(workers.items(), key=lambda item: len(item[1].n_projects), reverse=True)
        elif data == '5':
            workers_list = sorted(workers.items(), key=lambda item: item[1].n_success_days, reverse=True)
        else:
            print('\nДо свидания!')
            return '0'

        for i in range(len(workers_list)):
            print('{}. '
                  'Name = {}'
                  '\nn_lead = {}'
                  '\nn_success_delivery = {}'
                  '\nn_success_lead_days = {}'
                  '\nn_projects = {}'
                  '\nn_success_days = {}\n'
                  .format(i + 1,
                          workers_list[i][0],
                          len(workers_list[i][1].n_lead),
                          workers_list[i][1].n_success_delivery,
                          workers_list[i][1].n_success_lead_days,
                          len(workers_list[i][1].n_projects),
                          workers_list[i][1].n_success_days))
        return data
    except:
        print('\nПереход в главное меню...')
    return


if __name__ == '__main__':
    xlsx_dir = 'C:\\Users\\anton\\PycharmProjects\\excel_tensor\\xlsx_files_main\\'
    xlsx_path = get_xlsx_path(xlsx_dir)
    d_worker = get_names(xlsx_path)
    d_worker = get_info(xlsx_path, d_worker)

    while True:
        num = print_sorted(d_worker)
        if num == '0':
            break