class Row:
    COLOR = '46d444' # зеленый цвет
    def __new__(cls, name=None, *args, **kwargs):
        '''Создает объект класса только если передан аргумент - name.'''
        if name:
            return super().__new__(cls)

    
    def __init__(self, name):
        self.name = name # Номенклатура
        self.units_of_measurement = None # Ед. изм.
        self.initial_balance = None # Начальный остаток
        self.receipt_of_products = None # Приход
        self.sales = None # Продажи
        self.final_balance = None # Конечный остаток
        self.mark = None # Метка
        self.multiplicity = None # Кратность
        self.expiration_dates = None # Сроки годности
        self.declared = None # Заявлено | заяв
        self.diff = None # разн
        self.orders_is_on_the_way = {} # Заказы в пути {'29,11,': 50, '01,12,: 120}
        self.average_values = {}  # средние значения продаж => {'17,11,': 100, '24,11,': 111, ...}
        self.comments = None # комментарии


    def get_columns(self):
        return self.__dict__


    def __eq__(self, obj):
        if obj.__class__.__name__ == 'Row':
            return self.name == obj.name
        return self.name == obj

    def __lt__(self, obj):
        if obj.__class__.__name__ == 'Row':
            return self.name < obj.name
        return self.name < obj


    def __str__(self):
        return self.name


class Table:
    COLUMNS = {
        'name': ['Номенклатура', None, False],
        'units_of_measurement': ['Ед. изм.', None, False],
        'initial_balance': ['Начальный остаток', None, False],
        'receipt_of_products': ['Приход', None, False],
        'sales': ['Расход', None, True], # !!!!!!!!!!!!!!!!!!!!!!!!
        'final_balance': ['Конечный остаток', None, True],
        'mark': ['метка', None, False],
        'multiplicity': ['крат', None, False],
        'expiration_dates': ['сроки', None, False],
        'declared': ['заяв', None, True], # !!!!!!!!!!!!!!!!!!!!!!!!
        'diff': ['разн', None, True], # <<<<<<<<<<<<<<<<<<<
        'orders_is_on_the_way': ['заказ в пути', {}, True],
        'pud_order': ['пуд', None, True],
        'new_average_sales': ['ср нов', None, True],
        'krim_order': ['заказ', None, True],
        'remains': ['кон ост', None, False],
        'fact': ['факт', None, False],
        'medvedev_sales': ['медв', None, True],
        'tk_sales': ['тк', None, True],
        'atamanov_sales': ['атпр', None, True],
        'pud_sales': ['пудп', None, True],
        'average_values': ['ср', {}, True],
        'comments': ['комментарии', None, False],
        'sum': ['сум', None, True],
        'weight': ['вес', None, True]
    }
    START_HEADER = 3
    START_ROWS = 6
    END_ROWS = 500
    ORDERS_IS_ON_THE_WAY = {}
    AVERAGE_VALUES = {}
    
    def __init__(self):
        self.rows = []
        self.date = None


    def add_row(self, row):
        self.rows.append(row)


    def get_header(self):
        return self.COLUMNS


    def get_rows(self):
        '''
        Return sorted list self.rows
        '''
        return sorted(self.rows)


    def number_columns(self):
        num_col = 1
        for i in self.COLUMNS:
            if type(self.COLUMNS[i][1]) is dict:
                for j in self.COLUMNS[i][1]:
                    self.COLUMNS[i][1][j] = num_col
                    num_col += 1
            else:
                self.COLUMNS[i][1] = num_col
                num_col += 1



if __name__ == '__main__':
    r1 = Row('TEST')
    print(f"{r1.get_columns() = }") # получаем словарь аргументов и их значений
    r1.name = 'Obj'
    print(f"{r1 == 'Obj' = }") # сравнение на равенство объкта со строкой
    data = [r1]
    r2 = Row('TEST_2')
    data.append(r2)
    r2.name = 'Ok'
    r3 = Row('Test_3')
    r3.name = 'Obj'
    r4 = Row('TEST_4')
    r4.name = 'ABC'
    data.append(r4)
    print('data =', *data)
    print(f"{'Obj' in data = }") # определение вхождения
    print(f"{data.index('Obj') = }, {data.index(r1) = }") # нахождение индекса объекта по имени и самому объекту
    print(f"{r1 == r2 = }") # Сравнение объектов по имени
    print(f"{r1 == r3 = }") # Сравнение объектов по имени
    sorted_list = sorted(data)
    print(f"{sorted_list = }")
    print(*[(i, i.name) for i in sorted_list], sep='\n')
    ##################################################

    t = Table()
    t.get_header()
    t.add_row(r1)
    t.add_row(r2)
    t.add_row(r4)
    print(f"{t.get_rows() = }")
    print(f"{t.get_rows().index('ABC') = }")
    
