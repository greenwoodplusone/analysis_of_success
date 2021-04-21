import openpyxl
import glob
from Employee import *

class Operation:
    # Удаление переносов строки и пробелов после ФИО
    @staticmethod
    def remove_pass(value):
        index_new_line = value.index('\n')
        index_point = value.index('.')
        index_end_name = value.index('.', index_point + 1)
        if index_new_line == -1:
            if index_end_name == -1:
               return value
            else:
                return value[:index_end_name + 1]
        else:
            return value[:index_new_line]

    # Вычисляем успешность сотрудника по формуле:
    # успешность = (текущая успешность + план/факт) / 2 (2 - так как расчитываем средн. арифм. двух коэф.)
    #
    # value - Экземпляр класса Employee, plan - план, fact - факт
    @staticmethod
    def success_rate(value, plan, fact):

        if (type(plan) != int or plan == 0) and (type(fact) != int or fact == 0):
            # Сотрудник не участвовал в проекте, убавляем количество проектов у него
            value.count_project -= 1
        elif (type(plan) != int or plan == 0) and fact > 0 or plan == fact:
            # Сотрудник приступил к работе внепланово или выполнили работу в срок (факт = план)
            value.success_rate = (value.success_rate + 1) / 2
        elif plan > 0 and type(fact) != int or fact == 0:
            # Сотруднику дали работу, но он не приступил к ней, КПД сотрудника в этом проекте 0
            value.success_rate = (value.success_rate + 0) / 2
        else:
            # Базовое вычисление успешности
            value.success_rate = (value.success_rate + plan / fact) / 2

    # Перебор файлов, листов, сторк и колонок.
    # Добавление сотрудника в словарь, вычислеие и запись/изменение успешности при каждом переборе проектов.
    def calc_param(file_location):
        # Перебор каждого файла в папке
        # В файле берем только первый лист
        for file in glob.glob(file_location):
            book = openpyxl.open(file, read_only=True)
            sheet = book.active

            # номер колоки с которой произойдет считывание (счетчик начинается с 0)
            # задается в зависимости от расположения данных в ячейках (строках/столбцах) файла excel
            start_col = 4

            # перебираем колонки через одну, так как 1ая - план сотрудника, 2-ая - факт,
            # выходит что каждый новый сотрудник начниается на нечетной колонке (1, 3, ...)
            for col in range(start_col, sheet.max_column, 2):
                name_employee = Operation.remove_pass(sheet[1][col].value)

                # Добавляем сотрудника в словарь если его там нет
                if name_employee not in Employee.dict_employee:
                    Employee.dict_employee[name_employee] = Employee(name_employee)
                else:
                    continue

                # Перебор со 2ой и до последней строки файла, первая строка с данными ФИО
                for row in range(2, sheet.max_row + 1):
                    # Увеличиваем количество проектов на 1 для этого сотрудника
                    Employee.dict_employee[name_employee].count_project += 1

                    # Cокращение:
                    plan = sheet[row][col].value
                    fact = sheet[row][col + 1].value

                    # Добавляем в словарь количество смен план
                    if type(plan) == int and plan > 0:
                        Employee.dict_employee[name_employee].plan += plan

                    # Добавляем в словарь количество смен факт
                    if type(fact) == int and fact > 0:
                        Employee.dict_employee[name_employee].fact += fact

                    # Вычисление и запись успешности сотруднику
                    Operation.success_rate(Employee.dict_employee[name_employee], plan, fact)

    # Вывод отсортированного результата успешности сотрудников
    #
    # dict_employee - словарь с экземплярами класса
    @staticmethod
    def get_sorted_list(dict_employee):
        # Создание отсортированного списка только с именем сотрудника и его успешностью
        list_employee = []
        count = 0
        for employee in dict_employee:
            temp_list = [0, 0]
            temp_list[0] = format(dict_employee[employee].success_rate, '.3f')
            temp_list[1] = dict_employee[employee].name
            list_employee.append(temp_list)
            count += 1
        list_employee.sort(reverse=True)

        # Вывод в консоль
        count_employee = 1
        print('№  Успешность  Ф.И.О.')
        print('- - - - - - - - - - - - - - -')
        for success, employee in list_employee:
            print(count_employee, '. ', success, '   ', employee)
            count_employee += 1

        # Убрать документирование для создания текстового файла с результатами анализа
        """
        Operation.set_text_file(list_employee)
        """

    # Создание текстового файла и запись в него полученных данных
    @staticmethod
    def set_text_file(list_employee):
        my_file = open('Анализ успешности сотрудников.txt', 'w')
        count_employee = 1
        my_file.write('№  Успешность  Ф.И.О.\n')
        my_file.write('- - - - - - - - - - - - - - -\n')
        for item in list_employee:
            my_file.write(str(count_employee))
            my_file.write('.   ')
            my_file.write(str(item[0]))
            my_file.write('     ')
            my_file.write(str(item[1]))
            my_file.write('\n')
            count_employee += 1
        my_file.close()
