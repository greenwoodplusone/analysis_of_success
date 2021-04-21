# Создание сотрудника
#
# name - имя сотруника, count_project - количество проектов у сотрудника,
# plan - план, fact - факт, success_rate - коэффициент успешности
class Employee:
    # Хранит все экземпляры класса:
    dict_employee = {}

    def __init__(self, name, count_project=0, plan=0, fact=0, success_rate=0):
        self.__name = name
        self.__count_project = count_project
        self.__plan = plan
        self.__fact = fact
        self.__success_rate = success_rate

    @property
    def name(self):
        return self.__name

    @property
    def count_project(self):
        return self.__count_project

    @property
    def plan(self):
        return self.__plan

    @property
    def fact(self):
        return self.__fact

    @property
    def success_rate(self):
        return self.__success_rate

    @name.setter
    def name(self, value):
        if not isinstance(value, str):
            raise ValueError("Имя должно быть строкой")
        self.__name = value

    @count_project.setter
    def count_project(self, value):
        if not isinstance(value, (int, float)):
            raise ValueError("Количество проектов должно бысть числом")
        self.__count_project = value

    @plan.setter
    def plan(self, value):
        if not isinstance(value, (int, float)):
            raise ValueError("Количество проектов должно бысть числом")
        self.__plan = value

    @fact.setter
    def fact(self, value):
        if not isinstance(value, (int, float)):
            raise ValueError("Количество проектов должно бысть числом")
        self.__fact = value

    @success_rate.setter
    def success_rate(self, value):
        if not isinstance(value, (int, float)):
            raise ValueError("Количество проектов должно бысть числом")
        self.__success_rate = value
