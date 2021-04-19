import openpyxl
import glob

# Создание сотрудника
#
# name - имя сотруника, count_project - количество проектов у сотрудника, plan - план, fact - факт
class Employee:
    # Хранит все экземпляры класса:
    dict_employee = {}

    def __init__(self, name, count_project = 0, plan = 0, fact = 0, success_rate = 0):
        self.name = name
        self.count_project = count_project
        self.plan = plan
        self.fact = fact
        self.success_rate = success_rate

# Удаление переносов строки и пробелов после ФИО
def remov_pass (value):
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
def success_rate (value, plan, fact):
    if (type(plan) != int or plan == 0) and (type(fact) != int or fact == 0):
        # Сотрудник не участвовал в проекте, убавляем количество проектов у него
        value.count_project -= 1
    elif (type(plan) != int or plan == 0) and fact > 0:
        value.success_rate = (value.success_rate + 1) / 2
    elif plan == fact:
        value.success_rate = (value.success_rate + 1) / 2
    # Сотруднику дали работу, но он не приступил к ней, КПД сотрудника в этом проекте 0
    elif type(fact) != int or fact == 0:
        value.success_rate = (value.success_rate + 0) / 2
    else:
        value.success_rate = (value.success_rate + plan / fact) / 2

# Вывод отсортированного результата успешности сотрудников
#
# dict_employee - словарь с экземплярами класса
def get_sorted_list(dict_employee):
    # Создание списка только с именем сотрудника и его успешностью
    list = []
    count = 0
    for employee in dict_employee:
        temp_list = [0, 0]
        temp_list[0] = format(dict_employee[employee].success_rate, '.3f')
        temp_list[1] = dict_employee[employee].name
        list.append(temp_list)
        count += 1

    # Сортировка списка
    for i in range(0, len(list) - 1):
        for j in range(i + 1, len(list)):
            if list[j][0] > list[i][0]:
                list[i][0], list[j][0] = list[j][0], list[i][0]
                list[i][1], list[j][1] = list[j][1], list[i][1]

    # Вывод в консоль
    count_employee = 1
    print('№  Успешность  Ф.И.О.')
    print('- - - - - - - - - - - - - - -')
    for success, employee in list:
        print(count_employee, '. ', success, '   ', employee)
        count_employee += 1

# номер колоки с которой произойдет считывание (счетчик начинается с 0)
start_col = 4

for file in glob.glob("*.xlsx"):
    book = openpyxl.open(file, read_only=True)
    sheet = book.active

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
get_sorted_list (Employee.dict_employee)

"""
Убрать документирование для создания текстового файла с результатами анализа

# Создание текстового файла и запись в него полученных данных
my_file = open('Анализ успешности сотрудников.txt', 'w')
count_employee = 1
my_file.write('№  Успешность  Ф.И.О.\n')
my_file.write('- - - - - - - - - - - - - - -\n')
for item in list:
    my_file.write(str(count_employee))
    my_file.write('.   ')
    my_file.write(str(item[0]))
    my_file.write('     ')
    my_file.write(str(item[1]))
    my_file.write('\n')
    count_employee += 1
my_file.close()

"""