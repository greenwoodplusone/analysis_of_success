from Operation import *

# Расположение файлов
file_location = "*.xlsx"

# Вывод результата
if __name__ == "__main__":
    Operation.calc_param(file_location)
    Operation.get_sorted_list(Employee.dict_employee)
