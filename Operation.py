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