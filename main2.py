import os
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

    def __init__(self):
        self.n_lead = []
        self.n_success_delivery = 0
        self.n_success_lead_days = 0
        self.n_projects = []
        self.n_success_days = 0


class SuccessWorker:
    """
    __d_worker - словарь = { 'имя_сотрудника': Worker()}
    __xlsx_path - пути к .xlsx-файлам
    __workers_list - отсортированный список сотрудников по выбраному критерию
    -------------------------------------------------------------------------
    xlsx_dir - директория с excel-файлами
    data - критерий для оценки успешности
    -------------------------------------------------------------------------
    Критерии (в порядке степени важности):
    1: Составной критерий (по умолчанию),
    2: Количество руководящих должностей (быть руководителем проекта),
    3: Количество успешных сдач проектов в руководящей должности,
    4: Количество успешных дней в руководящей должности,
    5: Количество участий в проектах (за вычетом участий в проектах в руководящей должности),
    6: Количество успешных дней всего (за вычетом успешных дней в руководящей должности).
    -------------------------------------------------------------------------
    Составной критерий: вывод списка успешных сотрудников производится по степеням важности
    индивидуальных критериев (от наиболее важного (критерия номер 2) к наименее важному (критерию номер 6),
    т.е. сначала сравнивается 1 критерий, затем 2 критерий и т.д.).
    """
    def __init__(self):
        self.__d_worker = dict()
        self.__xlsx_path = []
        self.__workers_list = []

    def __get_xlsx_path(self, xlsx_dir):
        for file in os.listdir(xlsx_dir):
            if file.endswith('.xlsx'):
                self.__xlsx_path.append(os.path.join(xlsx_dir, file))

    def __get_names(self):
        for file in self.__xlsx_path:
            data = xlrd.open_workbook(filename=file)
            table = data.sheets()[0]

            # пробежаться по 0 строке и добавить в словарь все имена
            row_slice = table.row_slice(0, 4)
            for i in range(len(row_slice)):
                value = row_slice[i].value[:-6]
                if self.__d_worker.get(value) is None:
                    self.__d_worker[value] = Worker()

            # пробежаться по 1 столбцу и добавить в словарь все имена
            col_slice = table.col_slice(1, 1)
            for i in range(len(col_slice)):
                value = col_slice[i].value
                if self.__d_worker.get(value) is None:
                    self.__d_worker[value] = Worker()

    def __get_info(self):
        for file in self.__xlsx_path:
            data = xlrd.open_workbook(filename=file)
            table = data.sheets()[0]

            for row in range(1, table.nrows):
                for col in range(0, table.ncols, 2):
                    value = table.cell(row, col).value

                    if col == 0:
                        # подсчет проектов в руководящей должности
                        try:
                            if value not in self.__d_worker.get(table.cell(row, 1).value).n_lead:
                                self.__d_worker.get(table.cell(row, 1).value).n_lead.append(value)
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
                                self.__d_worker[table.cell(row, 1).value].n_success_delivery += 1
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

                            # если имя руководителя проекта не совпадает с именем участника проекта:
                            # успешные дни += days
                            if table.cell(0, col).value[:-6] != table.cell(row, 1).value:
                                self.__d_worker[table.cell(0, col).value[:-6]].n_success_days += days
                                # если названия проекта нет в списке проектов участника проекта: проекты += 1
                                if table.cell(row, 0).value not in self.__d_worker[table.cell(0, col).value[:-6]] \
                                        .n_projects:
                                    self.__d_worker[table.cell(0, col).value[:-6]] \
                                        .n_projects.append(table.cell(row, 0).value)
                            else:  # имя руководителя проекта совпадает с именем участника проекта
                                # успешные дни как руководителя проекта += days
                                self.__d_worker[table.cell(0, col).value[:-6]].n_success_lead_days += days
                                # если названия проекта нет в списке лидерских проектов: лидерские проекты += 1
                                if table.cell(row, 0).value not in self.__d_worker[
                                    table.cell(0, col).value[:-6]].n_lead:
                                    self.__d_worker[table.cell(0, col).value[:-6]].n_lead.append(
                                        table.cell(row, 0).value)
                        except:
                            print('Ошибка при подсчете успешных дней.')

    def __sorted_workers(self, data):
        if data == 1:
            self.__workers_list = sorted(self.__d_worker.items(), key=lambda item: (len(item[1].n_lead),
                                                                                    item[1].n_success_delivery,
                                                                                    item[1].n_success_lead_days,
                                                                                    len(item[1].n_projects),
                                                                                    item[1].n_success_days),
                                         reverse=True)
        elif data == 2:
            self.__workers_list = sorted(self.__d_worker.items(), key=lambda item: len(item[1].n_lead),
                                         reverse=True)
        elif data == 3:
            self.__workers_list = sorted(self.__d_worker.items(), key=lambda item: item[1].n_success_delivery,
                                         reverse=True)
        elif data == 4:
            self.__workers_list = sorted(self.__d_worker.items(), key=lambda item: item[1].n_success_lead_days,
                                         reverse=True)
        elif data == 5:
            self.__workers_list = sorted(self.__d_worker.items(), key=lambda item: len(item[1].n_projects),
                                         reverse=True)
        elif data == 6:
            self.__workers_list = sorted(self.__d_worker.items(), key=lambda item: item[1].n_success_days,
                                         reverse=True)
        else:
            raise ValueError('Имеется возможность выбирать критерии от 1 до 6.')
        return self.__workers_list

    def get_list_success_workers(self, xlsx_dir, data=1):
        """"""
        self.__get_xlsx_path(xlsx_dir)
        self.__get_names()
        self.__get_info()
        return self.__sorted_workers(data)


def print_list(workers_list):
    print()
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


if __name__ == '__main__':
    xlsx_dir = 'C:\\Users\\anton\\PycharmProjects\\excel_tensor\\xlsx_files_main\\'
    success_worker = SuccessWorker()
    print(success_worker.__doc__)
    workers = success_worker.get_list_success_workers(xlsx_dir)
    print_list(workers)