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

# --- НАЧАЛО КОДА ДЛЯ ЗАДАНИЯ 6 ---
if '6' in wb.sheetnames:
    ws_6 = wb['6']
    
    # 1. Применяем стили к шапке (A1:D1)
    for col_idx in range(1, 5):
        cell = ws_6.cell(row=1, column=col_idx)
        cell.fill = yellow_fill
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = thin_border

    # 2. Обработка данных
    for row in range(2, ws_6.max_row + 1):
        # Формула скидки: если цена < 7000, то 5%, иначе 7%
        ws_6[f'C{row}'].value = f'=IF(B{row}<7000, 0.05, 0.07)'
        # Формула цены со скидкой: Цена * (1 - Скидка)
        ws_6[f'D{row}'].value = f'=B{row}*(1-C{row})'
        
        # Форматирование
        ws_6[f'B{row}'].number_format = '#,##0 ₽' # Цена
        ws_6[f'C{row}'].number_format = '0%'      # Процент скидки
        ws_6[f'D{row}'].number_format = '#,##0 ₽' # Итог
        
        # Сетка для всей строки
        for col_idx in range(1, 5):
            ws_6.cell(row=row, column=col_idx).border = thin_border
            
    print("Задание 6 (лист 6) готово: скидки посчитаны, форматы настроены.")
# --- КОНЕЦ КОДА ДЛЯ ЗАДАНИЯ 6 ---


# --- НАЧАЛО КОДА ДЛЯ ЗАДАНИЯ 7 ---
if '7' in wb.sheetnames:
    ws_7 = wb['7']
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill

    # 1. Стили для шапки (A1:H1)
    for col_idx in range(1, 9):
        cell = ws_7.cell(row=1, column=col_idx)
        cell.fill = yellow_fill
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = thin_border

    # 2. Формулы и сетка
    for row in range(2, ws_7.max_row + 1):
        # Средний балл по 5 контрольным (B-F)
        ws_7[f'G{row}'].value = f'=AVERAGE(B{row}:F{row})'
        ws_7[f'G{row}'].number_format = '0.00'
        
        # Зачет/Незачет: если средний > 3, то зачет
        ws_7[f'H{row}'].value = f'=IF(G{row}>3, "зачет", "незачет")'
        
        for col_idx in range(1, 9):
            ws_7.cell(row=row, column=col_idx).border = thin_border

    # 3. Условное форматирование для столбца H (Зачет)
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

    # Если текст равен "незачет" -> красный
    ws_7.conditional_formatting.add('H2:H100', 
        CellIsRule(operator='equal', formula=['"незачет"'], fill=red_fill))
    # Если текст равен "зачет" -> зеленый
    ws_7.conditional_formatting.add('H2:H100', 
        CellIsRule(operator='equal', formula=['"зачет"'], fill=green_fill))

    # 4. Включаем автофильтр (чтобы ты мог нажать "Сортировка А-Я" по фамилии)
    ws_7.auto_filter.ref = ws_7.dimensions

    print("Задание 7 (лист 7) готово: средний балл, зачеты и авто-покраска настроены.")
# --- КОНЕЦ КОДА ДЛЯ ЗАДАНИЯ 7 ---

# --- НАЧАЛО КОДА ДЛЯ ЗАДАНИЯ 8 ---
if '8' in wb.sheetnames:
    ws_8 = wb['8']
    from openpyxl.styles import PatternFill, Font, Border, Side
    
    # 1. Настраиваем "зеленые" стили по образцу
    header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid") # Тёмно-зелёный для шапки
    data_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")   # Светло-зелёный для строк
    white_bold_font = Font(color="FFFFFF", bold=True) # Белый жирный шрифт для шапки
    black_font = Font(color="000000") # Обычный черный шрифт для данных
    
    # Тонкие границы (сеточка)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # 2. Оформляем шапку (1-я строка, столбцы A - E)
    for col_idx in range(1, 6):
        cell = ws_8.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = white_bold_font
        cell.border = thin_border

    # 3. Обрабатываем данные и красим строки (со 2-й строки до конца)
    for row in range(2, ws_8.max_row + 1):
        
        # Формула для столбца D (Выполнение плана):
        # Если продажи (C) > 1 000 000, то "выполнен", иначе "не выполнен"
        ws_8[f'D{row}'].value = f'=IF(C{row}>1000000, "выполнен", "не выполнен")'
        
        # Формула для столбца E (Зарплата):
        # Если план выполнен (D), то 20000 + 5% от продаж (C), иначе просто 20000
        ws_8[f'E{row}'].value = f'=IF(D{row}="выполнен", 20000+0.05*C{row}, 20000)'
        
        # Настраиваем денежный формат (рубли без копеек) для Продаж (C) и Зарплаты (E)
        currency_format = '#,##0 ₽'
        ws_8[f'C{row}'].number_format = currency_format
        ws_8[f'E{row}'].number_format = currency_format
        
        # Применяем светло-зеленую заливку и рамки для каждой ячейки в строке данных
        for col_idx in range(1, 6):
            cell = ws_8.cell(row=row, column=col_idx)
            cell.fill = data_fill
            cell.font = black_font
            cell.border = thin_border
            
    print("Задание 8 (лист 8) готово: выполнение плана посчитано, стили 'под зеленушку' применены.")
else:
    print("Внимание: Лист '8' не найден!")
# --- КОНЕЦ КОДА ДЛЯ ЗАДАНИЯ 8 ---

# Сохранение финального результата
output_filename = 'Практическая_работа_5.xlsx'
wb.save(output_filename)
print(f'\nВсе готово, бро! Файл "{output_filename}" успешно сохранен.')