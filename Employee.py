# Создание сотрудника
#
# name - имя сотруника, count_project - количество проектов у сотрудника,
# plan - план, fact - факт, success_rate - коэффициент успешности
class Employee:
    # Хранит все экземпляры класса:
    dict_employee = {}

    def __init__(self, name, count_project = 0, plan = 0, fact = 0, success_rate = 0):
        self.__name = name
        self.__count_project = count_project
        self.__plan = plan
        self.__fact = fact
        self.__success_rate = success_rate

    def get_name (self):
        return self.__name

    def get_count_project (self):
        return self.__count_project

    def get_plan (self):
        return self.__plan

    def get_fact (self):
        return self.__fact

    def get_success_rate (self):
        return self.__success_rate

    def set_name (self, value):
        if not isinstance(value, str):
            raise ValueError ("Имя должно быть строкой")
        self.__name = value

    def set_count_project (self, value):
        if not isinstance(value, (int, float)):
            raise ValueError ("Количество проектов должно бысть числом")
        self.__count_project = value

    def set_plan (self, value):
        if not isinstance(value, (int, float)):
            raise ValueError ("Количество проектов должно бысть числом")
        self.__plan = value

    def set_fact (self, value):
        if not isinstance(value, (int, float)):
            raise ValueError ("Количество проектов должно бысть числом")
        self.__fact = value

    def set_success_rate (self, value):
        if not isinstance(value, (int, float)):
            raise ValueError ("Количество проектов должно бысть числом")
        self.__success_rate = value

    name = property (fget=get_name, fset=set_name)
    count_project = property(fget=get_count_project, fset=set_count_project)
    plan = property(fget=get_plan, fset=set_plan)
    fact = property(fget=get_fact, fset=set_fact)
    success_rate = property(fget=get_success_rate, fset=set_success_rate)