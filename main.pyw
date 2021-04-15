import openpyxl
import glob

all_data = {}
count_employee = 0
start_col = 4 # номер колоки с которой произойдет считывание (счетчик начинается с 0)
for f in glob.glob("*.xlsx"):
    book = openpyxl.open(f, read_only=True)
    sheet = book.active
    for col in range(start_col, sheet.max_column, 2):
        name_employee = sheet[1][col].value

        # Добавляем сотрудника в словарь если его там нет
        if name_employee not in all_data:
            all_data[name_employee] = [0, 0, 0, 0, '']

            # Удаляем перенос строки и все лишнее после ФИО
            index_new_line = name_employee.index('\n')
            index_point = name_employee.index('.')
            index_end_name = name_employee.index('.', index_point + 1)
            if index_new_line == -1:
                if index_end_name == -1:
                    all_data[name_employee][start_col] = name_employee
                else:
                    all_data[name_employee][start_col] = name_employee[:index_end_name + 1]
            else:
                all_data[name_employee][start_col] = name_employee[:index_new_line]

            count_employee += 1

        for row in range(2, sheet.max_row + 1):
            # Увеличиваем количество проектов на 1 для этого сотрудника
            all_data[name_employee][2] += 1

            plan = sheet[row][col].value
            fact = sheet[row][col + 1].value

            # Добавляем в словарь количество смен план
            if type(plan) != int or plan < 0:
                all_data[name_employee][0] += 0
            else:
                all_data[name_employee][0] += plan

            # Добавляем в словарь количество смен факт
            if type(fact) != int or fact < 0:
                all_data[name_employee][1] += 0
            else:
                all_data[name_employee][1] += fact

            # Вычисляем успешность сотрудника по формуле:
            # успешность = (текущая успешность + план/факт) / 2 (2 - так как расчитываем средн. арифм. двух коэф.)
            if (type(plan) != int or plan == 0) and (type(fact) != int or fact == 0):
                # Сотрудник не участвовал в проекте, убавляем количесвто проектов у него
                all_data[name_employee][2] -= 1
            elif (type(plan) != int or plan == 0) and fact > 0:
                all_data[name_employee][3] = (all_data[name_employee][3] + 1) / 2
            elif plan == fact:
                all_data[name_employee][3] = (all_data[name_employee][3] + 1) / 2
            # Сотруднику дали работу, но он не приступил к ней, КПД сотрудника в этом проекте 0
            elif type(fact) != int or fact == 0:
                all_data[name_employee][3] = (all_data[name_employee][3] + 0) / 2
            else:
                all_data[name_employee][3] = (all_data[name_employee][3] + plan / fact) / 2

# Создание списка только с именем сотрудника и его успешностью
list = []
count = 0
for employee in all_data:
    temp_list = [0, 0]
    temp_list[0] = format(all_data[employee][3], '.3f')
    temp_list[1] = all_data[employee][4]
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
    print(count_employee, '. ',success, '   ',employee)
    count_employee += 1

"""
Убрать документирование для создания текстового файла с результатами анализа

# Создание текстового файла и запись в него полученных данных
my_file = open('Анализ успешности сотрудников.txt', 'w')
count_employee = 1
my_file.write('№  Успешность  Ф.И.О.\n')
my_file.write('- - - - - - - - - - - - - - -\n')
for i in list:
    my_file.write(str(count_employee))
    my_file.write('.   ')
    my_file.write(str(i[0]))
    my_file.write('     ')
    my_file.write(str(i[1]))
    my_file.write('\n')
    count_employee += 1
my_file.close()

"""