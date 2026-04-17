import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

# Импорты для диаграмм (ДОБАВИЛ ScatterChart и Series)
from openpyxl.chart import BarChart, PieChart3D, ScatterChart, Reference, Series
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

cell_a1 = cast(Cell, ws1['A1'])
cell_a1.value = 'Показатели производства за 2014 год'
ws1.merge_cells('A1:E1')
cell_a1.font = Font(bold=True)
cell_a1.alignment = Alignment(horizontal='center', vertical='center')

headers_ws1 = ['1', '2', '3', '4']
for col_idx, text in enumerate(headers_ws1, start=2):
    c = cast(Cell, ws1.cell(row=2, column=col_idx))
    c.value = text
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal='center')

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

headers_ws2 = ['1 квартал', '2 квартал', '3 квартал', '4 квартал']
for col_idx, text in enumerate(headers_ws2, start=2):
    c = cast(Cell, ws2.cell(row=1, column=col_idx))
    c.value = text
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal='center')

data_ws2 = ['Факт (тыс.руб.)', 980, 1150, 1200, 1060]
for col_idx, val in enumerate(data_ws2, start=1):
    c = cast(Cell, ws2.cell(row=2, column=col_idx))
    c.value = val
    if col_idx == 1: c.font = Font(bold=True)

ws2.column_dimensions['A'].width = 18
for col in ['B', 'C', 'D', 'E']:
    ws2.column_dimensions[col].width = 12

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
# ЗАДАНИЯ 3, 4, 5, 6: МАТЕМАТИЧЕСКИЕ ФОРМУЛЫ
# ==========================================
ws3 = cast(Worksheet, wb.create_sheet(title="Формулы (4-6)"))

ws3['A1'] = 'X'
ws3['B1'] = 'Y'
ws3['A2'] = 4  
ws3['B2'] = 3  

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

ws3['D1'] = '=(1+A2)/(4*B2)'
ws3['D2'] = '=-2*A2+(A2^5)/(3*B2^2+4)'
ws3['D3'] = '=SQRT(7*A2+2)'
ws3['D4'] = '=SIN((A2+5)/(3*A2-2))+SQRT(A2^3+1)'

for row_idx in range(1, 5):
    res_cell = cast(Cell, ws3.cell(row=row_idx, column=4))
    res_cell.number_format = '0.000'
    res_cell.alignment = Alignment(horizontal='left')

for cell_ref in ['A1', 'B1']:
    c = cast(Cell, ws3[cell_ref])
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal='center')
    c.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

ws3.column_dimensions['C'].width = 12
ws3.column_dimensions['D'].width = 12


# ==========================================
# ЗАДАНИЕ 7: ПОСТРОЕНИЕ ГРАФИКА ФУНКЦИИ
# ==========================================
ws7 = cast(Worksheet, wb.create_sheet(title="Задание 7"))

# Шапка таблицы
ws7['A1'] = 'X'
ws7['B1'] = 'Y'

for c_ref in ['A1', 'B1']:
    c = cast(Cell, ws7[c_ref])
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal='center')

# Заполняем таблицу с помощью цикла while (от -4 до 4 с шагом 0.2)
current_x = -4.0
current_row = 2

# Используем 4.01, чтобы избежать багов с дробными числами (когда 4.0 превращается в 4.00000001)
while current_x <= 4.01:
    # Записываем значение X
    cell_x = cast(Cell, ws7.cell(row=current_row, column=1))
    cell_x.value = current_x
    cell_x.number_format = '0.00' # 2 знака после запятой
    
    # Записываем Excel-формулу для Y. Важно: используем точку (2.5), так как это стандартный формат формул
    cell_y = cast(Cell, ws7.cell(row=current_row, column=2))
    cell_y.value = f'=SIN(2.5*(A{current_row}-3))'
    cell_y.number_format = '0.00'
    
    current_x += 0.2
    current_row += 1

# Строим точечную диаграмму
chart7 = ScatterChart()
chart7.title = "y=sin 2.5(x-3)"
chart7.style = 2

# Настраиваем жесткие границы осей как в PDF
chart7.x_axis.scaling.min = -5.0
chart7.x_axis.scaling.max = 5.0
chart7.y_axis.scaling.min = -2.0
chart7.y_axis.scaling.max = 2.0
chart7.y_axis.majorUnit = 1.0 # Основные деления оси Y через 1.0

# Передаем данные в график
x_values = Reference(ws7, min_col=1, min_row=2, max_row=current_row-1)
y_values = Reference(ws7, min_col=2, min_row=2, max_row=current_row-1)

# Создаем ряд данных (график)
series7 = Series(y_values, x_values, title_from_data=False)

# УДАЛИЛИ ПРОБЛЕМНУЮ СТРОЧКУ ЗДЕСЬ

# МАГИЯ: Делаем кривую гладкой
series7.smooth = True 

chart7.series.append(series7)

# Размещаем график рядышком с таблицей (в колонке D)
ws7.add_chart(chart7, "D2")

# ==========================================
# ФИНАЛЬНОЕ СОХРАНЕНИЕ
# ==========================================
filename = 'Практическая работа 4 (Полная).xlsx'
wb.save(filename)
print(f'Отлично, бро! Файл "{filename}" обновлен. График функции построен!')