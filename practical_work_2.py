# Импортируем библиотеку для работы с Excel
import openpyxl
# Импортируем инструменты для настройки внешнего вида ячеек
from openpyxl.styles import Alignment, Font, PatternFill

# Импортируем типы данных, чтобы подсказать Pyright, с чем мы работаем
from typing import cast
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell

# 1. Создаем новую рабочую книгу (файл)
wb = openpyxl.Workbook()

# Успокаиваем Pyright: жестко указываем, что активный лист — это объект Worksheet
ws = cast(Worksheet, wb.active)
ws.title = "Практическая работа 2"

# 2. Оформляем главный заголовок таблицы
# Pyright сомневается, одна это ячейка или диапазон, поэтому применяем cast(Cell, ...)
cell_a1 = cast(Cell, ws['A1'])
cell_a1.value = 'Анализ спроса и продаж продукции фирмы "Шанс"'

ws.merge_cells('A1:H1') # Объединяем ячейки от A до H
cell_a1.font = Font(bold=True, size=12)
cell_a1.alignment = Alignment(horizontal='center', vertical='center')

# 3. Настраиваем шапку таблицы (названия столбцов)
headers = [
    ('A2', 'Наименование\nпродукции', 'A2:A3'),
    ('B2', 'Цена за\nед.', 'B2:B3'),
    ('C2', 'Спрос,\nшт.', 'C2:C3'),
    ('D2', 'Предл.,\nшт.', 'D2:D3'),
    ('E2', 'Продажа', 'E2:G2'), 
    ('H2', 'Выручка\nот\nпродаж', 'H2:H3')
]

for cell_ref, text, merge_range in headers:
    current_cell = cast(Cell, ws[cell_ref])
    current_cell.value = text
    ws.merge_cells(merge_range)

# Добавляем подзаголовки
cast(Cell, ws['E3']).value = 'Безнал.'
cast(Cell, ws['F3']).value = 'Нал.'
cast(Cell, ws['G3']).value = 'Всего'

# Применяем форматирование ко всем ячейкам шапки
header_fill = PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid")

for row in ws['A2:H3']:
    for cell in row:
        # В циклах по строкам получаем ячейки, тоже подсказываем их тип
        c = cast(Cell, cell)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.font = Font(bold=True)
        c.fill = header_fill

# 4. Вносим исходные данные
data = [
    ['Телевизоры', 350.35, 13, 15, 5, 7],
    ['Видеомагнитофоны', 320.00, 70, 65, 30, 35],
    ['Проигрыватели', 400.21, 65, 134, 40, 26]
]

# Метод .cell() более безопасен для типов, так что тут Pyright ругаться не должен
for i, row_data in enumerate(data, start=4):
    ws.cell(row=i, column=1, value=row_data[0]) 
    ws.cell(row=i, column=2, value=row_data[1]) 
    ws.cell(row=i, column=3, value=row_data[2]) 
    ws.cell(row=i, column=4, value=row_data[3]) 
    ws.cell(row=i, column=5, value=row_data[4]) 
    ws.cell(row=i, column=6, value=row_data[5]) 
    
    # 5. Прописываем формулы Excel
    ws.cell(row=i, column=7, value=f'=E{i}+F{i}')
    ws.cell(row=i, column=8, value=f'=B{i}*F{i}')

# 6. Применяем денежный формат
for row_num in range(4, 7):
    cast(Cell, ws[f'B{row_num}']).number_format = '$ #,##0.00'
    cast(Cell, ws[f'H{row_num}']).number_format = '$ #,##0.00'

# Настраиваем ширину столбцов
ws.column_dimensions['A'].width = 20
ws.column_dimensions['B'].width = 12
ws.column_dimensions['E'].width = 10
ws.column_dimensions['F'].width = 10
ws.column_dimensions['G'].width = 10
ws.column_dimensions['H'].width = 15

# 7. Сохраняем готовый файл
filename = 'Практическая работа 2.xlsx'
wb.save(filename)
print(f'Готово! Файл "{filename}" успешно создан, а Pyright теперь счастлив.')