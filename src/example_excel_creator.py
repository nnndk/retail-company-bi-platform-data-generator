from src.excel_generator.excel_fact_dimension import ExcelFactDimension
from src.excel_generator.excel_cube import ExcelCube

from datetime import datetime


def create_example_excel():
    """
    This method shows the way of creating your own Excel test file.
    It lets to set dimensions, and periods. Also, you can configure limitations of value generation
    :return:
    """
    customers = {'value': ['Смирнов', 'Иванов', 'Кузнецов', 'Соколов', 'Попов', 'Лебедев', 'Козлов', 'Новиков']}
    dim_customer = ExcelCube.create_excel_dim('Покупатель', ['value'], customers)

    shops = {'value': ['Магазин 1', 'Магазин 2', 'Магазин 3', 'Магазин 4']}
    dim_shop = ExcelCube.create_excel_dim('Магазин', ['value'], shops)

    goods = {'value': ['Карандаш', 'Пластелин', 'Ручка', 'Тетрадь', 'Циркуль', 'Калькулятор'],
             'Цена': ['25', '210', '35', '12', '300', '500']}
    dim_goods = ExcelCube.create_excel_dim('Товар', ['Название', 'Цена'], goods)

    facts = [ExcelFactDimension('Количество', 23, 55)]

    cube = ExcelCube('Данные о продажах', [dim_customer, dim_shop, dim_goods], facts)
    start_date, end_date = datetime(2022, 1, 1), datetime(2023, 12, 31)
    cube.generate(start_date, end_date, 3, 7, 'sales_data')
