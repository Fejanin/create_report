from my_lib.base_func import create_base_table_from_dv_file
from my_lib.base_func import add_data_from_old_file
from my_lib.base_func import add_data_from_sales_file
from my_lib.base_func import create_new_file


# test files
file_dv = 'дв 06,02,24ост.xlsx'
file_dv_old = ''
file_sales = ''
file_declared = ''


# за основу берем файл движения
table = create_base_table_from_dv_file(file_dv)

'''
# извлекаем данные из файла прошлой недели и добавляем данные в таблицу
add_data_from_old_file(file_dv_old, table)


# извлекаем данные из файла продаж
add_data_from_sales_file(file_sales, table, ('количество', 'sales'))


# извлекаем данные из файла заявлено
add_data_from_sales_file(file_declared, table, ('Отгружено', 'declared'))
'''
# создаем новый файл и заносим в него данные
create_new_file(f'new_{file_dv}', table)


for i in table.get_rows():
    print(i.get_columns())