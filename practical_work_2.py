# Импортируем библиотеку для работы с Excel
import openpyxl
# Импортируем инструменты для настройки внешнего вида ячеек (выравнивание, шрифт)
from openpyxl.styles import Alignment, Font, PatternFill

# 1. Создаем новую рабочую книгу (файл) и выбираем активный лист
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Практическая работа 2"

# 2. Оформляем главный заголовок таблицы
ws['A1'] = 'Анализ спроса и продаж продукции фирмы "Шанс"'
ws.merge_cells('A1:H1') # Объединяем ячейки от A до H
# Делаем текст жирным и выравниваем по центру
ws['A1'].font = Font(bold=True, size=12)
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

# 3. Настраиваем шапку таблицы (названия столбцов)
# Список с настройками: (Ячейка, Текст, Диапазон для объединения)
headers = [
    ('A2', 'Наименование\nпродукции', 'A2:A3'),
    ('B2', 'Цена за\nед.', 'B2:B3'),
    ('C2', 'Спрос,\nшт.', 'C2:C3'),
    ('D2', 'Предл.,\nшт.', 'D2:D3'),
    ('E2', 'Продажа', 'E2:G2'), # Объединяем 3 ячейки по горизонтали
    ('H2', 'Выручка\nот\nпродаж', 'H2:H3')
]

# Проходимся по списку и применяем настройки
for cell, text, merge_range in headers:
    ws[cell] = text
    ws.merge_cells(merge_range)

# Добавляем подзаголовки для столбца "Продажа"
ws['E3'] = 'Безнал.'
ws['F3'] = 'Нал.'
ws['G3'] = 'Всего'

# Применяем форматирование ко всем ячейкам шапки (выравнивание по центру и перенос слов)
# Для красоты также добавим легкий фиолетовый фон, как на скриншоте (по желанию)
header_fill = PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid")

for row in ws['A2:H3']:
    for cell in row:
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.font = Font(bold=True)
        cell.fill = header_fill

# 4. Вносим исходные данные
data = [
    ['Телевизоры', 350.35, 13, 15, 5, 7],
    ['Видеомагнитофоны', 320.00, 70, 65, 30, 35],
    ['Проигрыватели', 400.21, 65, 134, 40, 26]
]

# Записываем данные начиная с 4-й строки
for i, row_data in enumerate(data, start=4):
    ws.cell(row=i, column=1, value=row_data[0]) # Наименование
    ws.cell(row=i, column=2, value=row_data[1]) # Цена
    ws.cell(row=i, column=3, value=row_data[2]) # Спрос
    ws.cell(row=i, column=4, value=row_data[3]) # Предложение
    ws.cell(row=i, column=5, value=row_data[4]) # Безнал
    ws.cell(row=i, column=6, value=row_data[5]) # Нал
    
    # 5. Прописываем формулы Excel прямо через Python
    # Формула «Всего» = «Безнал» (E) + «Нал» (F)
    ws.cell(row=i, column=7, value=f'=E{i}+F{i}')
    
    # Формула «Выручка» = «Цена за ед.» (B) * «Нал.» (F)
    ws.cell(row=i, column=8, value=f'=B{i}*F{i}')

# 6. Применяем денежный формат для столбцов с ценой и выручкой
# Формат '$ #,##0.00' означает: знак доллара, разделитель тысяч (если будет) и два знака после запятой
for row_num in range(4, 7):
    ws[f'B{row_num}'].number_format = '$ #,##0.00'
    ws[f'H{row_num}'].number_format = '$ #,##0.00'

# Немного расширим столбцы, чтобы всё красиво влезло
ws.column_dimensions['A'].width = 20
ws.column_dimensions['B'].width = 12
ws.column_dimensions['E'].width = 10
ws.column_dimensions['F'].width = 10
ws.column_dimensions['G'].width = 10
ws.column_dimensions['H'].width = 15

# 7. Сохраняем готовый файл
filename = 'Практическая работа 2.xlsx'
wb.save(filename)
print(f'Готово! Файл "{filename}" успешно создан в папке со скриптом.')