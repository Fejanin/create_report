from my_lib.base_func import create_base_table_from_dv_file
from my_lib.base_func import add_data_from_old_file
from my_lib.base_func import add_data_from_sales_file
from my_lib.base_func import create_new_file
import os
import re


simf_pat = {
    'движение': r'(дв ([0-9]+,)+[0-9]+[а-яё]+\.xlsx)',
    'продажи': r'(заяв ([0-9]+)(,[0-9]+)?-([0-9]+,){2}[0-9]+\.xlsx)',
    'заявлено': r'(пр ([0-9]+)(,[0-9]+)?-([0-9]+,){2}[0-9]+\.xlsx)',
}

# TODO создать подменю выбора склада и присвоить соответсвующие паттерны в patterns
patterns = simf_pat

# получаем список файлов для обработки
all_files_list = ' '.join([i for i in os.listdir() if '.xlsx' in i])
groups_of_files = {
    'движение': [i[0] for i in re.findall(patterns['движение'], all_files_list)],
    'продажи': [i[0] for i in re.findall(patterns['продажи'], all_files_list)],
    'заявлено': [i[0] for i in re.findall(patterns['заявлено'], all_files_list)]
}

print(groups_of_files)

# test files
file_dv = 'дв 05,01,24пок.xlsx'
file_dv_old = 'дв 03,01,24пок.xlsx'
file_sales = 'пр 29-05,01,24.xlsx'
file_declared = 'заяв 29-05,01,24.xlsx'


# за основу берем файл движения
pattern_current_data = r'([0-9]{2},){2}'
current_data = re.search(pattern_current_data, file_dv)[0]
# print(f'{current_data = }')
table = create_base_table_from_dv_file(file_dv)


# извлекаем данные из файла прошлой недели и добавляем данные в таблицу
add_data_from_old_file(file_dv_old, table)


# извлекаем данные из файла продаж
add_data_from_sales_file(file_sales, table, ('количество', 'sales'))


# извлекаем данные из файла заявлено
add_data_from_sales_file(file_declared, table, ('Отгружено', 'declared'))

# создаем новый файл и заносим в него данные
create_new_file(f'new_{file_dv}', table)
