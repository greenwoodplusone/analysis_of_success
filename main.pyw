import openpyxl
import glob
from Employee import *
from Operation import *

def main():
    # Перебор каждого файла в папке
    # В файле берем только первый лист
    for file in glob.glob("*.xlsx"):
        book = openpyxl.open(file, read_only=True)
        sheet = book.active

        # номер колоки с которой произойдет считывание (счетчик начинается с 0)
        # задается в зависимости от расположения данных в ячейках (строках/столбцах) файла excel
        start_col = 4

        # перебираем колонки через одну, так как 1ая - план сотрудника, 2-ая - факт,
        # выходит что каждый новый сотрудник начниается на нечетной колонке (1, 3, ...)
        for col in range(start_col, sheet.max_column, 2):
            name_employee = remov_pass(sheet[1][col].value)

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
                success_rate(Employee.dict_employee[name_employee], plan, fact)

# Вывод результата
if __name__ == "__main__":
    main()
    get_sorted_list (Employee.dict_employee)

"""
Убрать документирование для создания текстового файла с результатами анализа

set_text_file(list)
"""