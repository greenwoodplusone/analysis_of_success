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

# Создание текстового файла и запись в него полученных данных
def set_text_file(list):
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