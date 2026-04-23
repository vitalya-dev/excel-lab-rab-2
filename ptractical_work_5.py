import openpyxl

# Загрузка существующего файла
try:
    wb = openpyxl.load_workbook("Фильтрация данных.xlsx")
    print("Исходный файл успешно загружен!")
except FileNotFoundError:
    print("Ошибка: Файл 'Фильтрация данных.xlsx' не найден. Убедись, что он лежит в той же папке.")
    exit()

# --- НАЧАЛО КОДА ДЛЯ ЗАДАНИЯ 1 ---
if '1_1' in wb.sheetnames:
    ws_1_1 = wb['1_1']
    # Включаем кнопочки фильтра для всей таблицы
    ws_1_1.auto_filter.ref = ws_1_1.dimensions
    print("Задание 1 (лист 1_1): Кнопки фильтра добавлены.")

if '1_2' in wb.sheetnames:
    ws_1_2 = wb['1_2']
    ws_1_2.auto_filter.ref = ws_1_2.dimensions
    print("Задание 1 (лист 1_2): Кнопки фильтра добавлены.")
# --- КОНЕЦ КОДА ДЛЯ ЗАДАНИЯ 1 ---

# --- НАЧАЛО КОДА ДЛЯ ЗАДАНИЯ 2 ---
if '2' in wb.sheetnames:
    ws_2 = wb['2']
    # Включаем фильтр для листа 2
    ws_2.auto_filter.ref = ws_2.dimensions
    print("Задание 2 (лист 2): Кнопки фильтра добавлены. (В Excel выбери Числовые фильтры -> Первые 10 -> укажи 12)")
# --- КОНЕЦ КОДА ДЛЯ ЗАДАНИЯ 2 ---

# --- НАЧАЛО КОДА ДЛЯ ЗАДАНИЯ 3 ---
if '3' in wb.sheetnames:
    ws_3 = wb['3']
    # Включаем фильтр для листа 3
    ws_3.auto_filter.ref = ws_3.dimensions
    print("Задание 3 (лист 3): Кнопки фильтра добавлены. (В Excel выбери Числовые фильтры -> Больше -> 25000)")
# --- КОНЕЦ КОДА ДЛЯ ЗАДАНИЯ 3 ---

# --- НАЧАЛО КОДА ДЛЯ ЗАДАНИЯ 4 ---
if '4' in wb.sheetnames:
    ws_4 = wb['4']
    # Включаем фильтр для листа 4, чтобы появилась возможность сортировки
    ws_4.auto_filter.ref = ws_4.dimensions
    print("Задание 4 (лист 4): Кнопки фильтра добавлены. (В Excel отсортируй по столбцу 'Торговый представитель' от А до Я)")
# --- КОНЕЦ КОДА ДЛЯ ЗАДАНИЯ 4 ---

# Если в начале файла еще нет этих импортов, обязательно добавь их:
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# --- НАЧАЛО КОДА ДЛЯ ЗАДАНИЯ 5 ---
if '5' in wb.sheetnames:
    ws_5 = wb['5']
    
    # 1. Настраиваем стили: желтенькая заливка, жирный шрифт, выравнивание по центру и тонкие границы
    yellow_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid") # Приятный желтый цвет
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # 2. Красим шапку таблицы (1-я строка, столбцы с A до E)
    for col_idx in range(1, 6): # Столбцы от 1 (A) до 5 (E)
        header_cell = ws_5.cell(row=1, column=col_idx)
        header_cell.fill = yellow_fill
        header_cell.font = bold_font
        header_cell.alignment = center_align
        header_cell.border = thin_border
    
    # 3. Проходимся по всем строкам данных (со 2-й до конца)
    for row in range(2, ws_5.max_row + 1):
        
        # Записываем формулы
        ws_5[f'D{row}'].value = f'=IF(B{row}>10, 2, 1)'
        ws_5[f'E{row}'].value = f'=D{row}*C{row}'
        
        # Денежный формат
        currency_format = '#,##0 ₽'
        ws_5[f'C{row}'].number_format = currency_format
        ws_5[f'E{row}'].number_format = currency_format
        
        # Добавляем границы (сеточку) для каждой ячейки в строке, как на скрине
        for col_idx in range(1, 6):
            ws_5.cell(row=row, column=col_idx).border = thin_border
            
    print("Задание 5 (лист 5): Формулы, форматы и желтенькая шапка с границами успешно добавлены!")
else:
    print("Внимание: Лист '5' не найден!")
# --- КОНЕЦ КОДА ДЛЯ ЗАДАНИЯ 5 ---

# Сохранение финального результата
output_filename = 'Практическая_работа_5.xlsx'
wb.save(output_filename)
print(f'\nВсе готово, бро! Файл "{output_filename}" успешно сохранен.')