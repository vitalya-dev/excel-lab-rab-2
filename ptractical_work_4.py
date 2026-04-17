import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.chart import BarChart, PieChart3D, Reference
from openpyxl.chart.label import DataLabelList

# Типизация для Pyright
from typing import cast
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell

# Создаем рабочую книгу
wb = openpyxl.Workbook()

# ==========================================
# ЗАДАНИЕ 1: ГИСТОГРАММА ПЛАН/ФАКТ
# ==========================================
ws1 = cast(Worksheet, wb.active)
ws1.title = "Задание 1"

# Заголовок
cell_a1 = cast(Cell, ws1['A1'])
cell_a1.value = 'Показатели производства за 2014 год'
ws1.merge_cells('A1:E1')
cell_a1.font = Font(bold=True)
cell_a1.alignment = Alignment(horizontal='center', vertical='center')

# Шапка кварталов
headers_ws1 = ['1', '2', '3', '4']
for col_idx, text in enumerate(headers_ws1, start=2):
    c = cast(Cell, ws1.cell(row=2, column=col_idx))
    c.value = text
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal='center')

# Данные
data_ws1 = [
    ['План (тыс. руб)', 1000, 800, 1500, 1100],
    ['Факт (тыс. руб)', 980, 1150, 1200, 1060]
]

for i, row_data in enumerate(data_ws1, start=3):
    for j, val in enumerate(row_data, start=1):
        cell = cast(Cell, ws1.cell(row=i, column=j))
        cell.value = val
        if j == 1: cell.font = Font(bold=True)
        else: cell.alignment = Alignment(horizontal='center')

ws1.column_dimensions['A'].width = 18

# Построение гистограммы
chart1 = BarChart()
chart1.type = "col"
chart1.title = "Поквартальное выполнение плана"
data_ref1 = Reference(ws1, min_col=1, min_row=3, max_col=5, max_row=4)
cats_ref1 = Reference(ws1, min_col=2, min_row=2, max_col=5, max_row=2)
chart1.add_data(data_ref1, titles_from_data=True, from_rows=True)
chart1.set_categories(cats_ref1)
ws1.add_chart(chart1, "A6")


# ==========================================
# ЗАДАНИЕ 2: ОБЪЕМНАЯ КРУГОВАЯ ДИАГРАММА
# ==========================================
ws2 = cast(Worksheet, wb.create_sheet(title="Задание 2"))

# Шапка и данные
headers_ws2 = ['1 квартал', '2 квартал', '3 квартал', '4 квартал']
for col_idx, text in enumerate(headers_ws2, start=2):
    c = cast(Cell, ws2.cell(row=1, column=col_idx))
    c.value = text
    c.font = Font(bold=True)

data_ws2 = ['Факт (тыс.руб.)', 980, 1150, 1200, 1060]
for col_idx, val in enumerate(data_ws2, start=1):
    c = cast(Cell, ws2.cell(row=2, column=col_idx))
    c.value = val
    if col_idx == 1: c.font = Font(bold=True)

ws2.column_dimensions['A'].width = 18

# Построение 3D круговой диаграммы
pie = PieChart3D()
pie.title = "Фактическое выполнение плана"
data_ref2 = Reference(ws2, min_col=2, min_row=2, max_col=5, max_row=2)
cats_ref2 = Reference(ws2, min_col=2, min_row=1, max_col=5, max_row=1)
pie.add_data(data_ref2, from_rows=True)
pie.set_categories(cats_ref2)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
ws2.add_chart(pie, "A4")


# ==========================================
# ЗАДАНИЯ 3, 4, 5: МАТЕМАТИЧЕСКИЕ ФОРМУЛЫ
# ==========================================
ws3 = cast(Worksheet, wb.create_sheet(title="Формулы (4-6)"))

# 1. Создаем структуру таблицы 
ws3['A1'] = 'X'
ws3['B1'] = 'Y'
ws3['A2'] = 4  # Значение x из примера [cite: 158, 166]
ws3['B2'] = 3  # Значение y из примера [cite: 158, 166]

# Подписи для заданий в колонке C
tasks = [
    ('C1', 'Задание 3'),
    ('C2', 'Задание 4'),
    ('C3', 'Задание 5'),
    ('C4', 'Задание 6')
]

for cell_ref, label in tasks:
    c = cast(Cell, ws3[cell_ref])
    c.value = label
    c.font = Font(bold=True)

# 2. Прописываем формулы в колонку D
# Задание 3: (1+x)/(4y) 
ws3['D1'] = '=(1+A2)/(4*B2)'

# Задание 4: -2x + (x^5)/(3y^2+4) 
ws3['D2'] = '=-2*A2+(A2^5)/(3*B2^2+4)'

# Задание 5: КОРЕНЬ(7x+2) [cite: 177, 179]
ws3['D3'] = '=SQRT(7*A2+2)'

# Задание 6: sin((x+5)/(3x-2)) + КОРЕНЬ(x^3+1) [cite: 201]
ws3['D4'] = '=SIN((A2+5)/(3*A2-2))+SQRT(A2^3+1)'

# 3. Настраиваем формат: 3 знака после запятой для результатов 
for row_idx in range(1, 5):
    res_cell = cast(Cell, ws3.cell(row=row_idx, column=4))
    res_cell.number_format = '0.000'
    res_cell.alignment = Alignment(horizontal='left')

# Красивое оформление шапки X и Y
for cell_ref in ['A1', 'B1']:
    c = cast(Cell, ws3[cell_ref])
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal='center')
    c.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

ws3.column_dimensions['C'].width = 12
ws3.column_dimensions['D'].width = 12

# Сохранение
filename = 'Практическая работа 4 (Полная).xlsx'
wb.save(filename)
print(f'Бро, готово! Собрал всё в файл "{filename}". Все диаграммы и формулы на месте!')