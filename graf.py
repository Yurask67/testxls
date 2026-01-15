#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ПРОФЕССИОНАЛЬНЫЙ ГЕНЕРАТОР ГРАФИКА ОТПУСКОВ 2026
С УЧЁТОМ РАСШИРЕНИЯ И ПУСТЫХ СТРОК
"""

import os
import sys
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==================== НАСТРОЙКИ ====================
MAX_EMPLOYEES = 20  # Максимальное количество сотрудников (с запасом)
COMPANY_NAME = "НАЗВАНИЕ КОМПАНИИ"  # Название компании для заголовков
YEAR = 2026  # Год графика
# ==================================================

def get_russian_calendar(year=YEAR):
    """Возвращает производственный календарь России на указанный год"""
    
    # Праздничные дни (нерабочие) для 2026 года
    holidays_2026 = [
        # Новогодние каникулы и Рождество
        (year, 1, 1), (year, 1, 2), (year, 1, 3), (year, 1, 4),
        (year, 1, 5), (year, 1, 6), (year, 1, 7), (year, 1, 8),
        (year, 1, 9),  # 9 января - дополнительный выходной
        
        # 23 Февраля
        (year, 2, 23),
        
        # 8 Марта
        (year, 3, 8),
        
        # 1 Мая
        (year, 5, 1),
        
        # 9 Мая
        (year, 5, 9),
        
        # 12 Июня
        (year, 6, 12),
        
        # 4 Ноября
        (year, 11, 4),
    ]
    
    # Предпраздничные дни (сокращенные на 1 час) для 2026
    pre_holidays_2026 = [
        (year, 2, 20),  # Пятница перед 23 февраля
        (year, 3, 7),   # Суббота перед 8 марта (рабочая)
        (year, 5, 8),   # Пятница перед 9 мая
        (year, 6, 11),  # Пятница перед 12 июня
        (year, 11, 3),  # Вторник перед 4 ноября
        (year, 12, 31), # Четверг перед Новым годом
    ]
    
    # Рабочие субботы (переносы) для 2026
    working_saturdays_2026 = [
        (year, 2, 21),  # Суббота (рабочая вместо понедельника)
        (year, 11, 14), # Суббота (рабочая вместо понедельника)
    ]
    
    # Определяем, високосный ли год
    is_leap = (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)
    days_in_year = 366 if is_leap else 365
    
    # Создаем календарь на весь год
    calendar = {}
    start_date = datetime(year, 1, 1)
    
    for i in range(days_in_year):
        current_date = start_date + timedelta(days=i)
        if current_date.year > year:
            break
            
        date_key = current_date.date()
        weekday = current_date.weekday()  # 0=пн, 6=вс
        
        # Определяем тип дня
        is_holiday = (current_date.year, current_date.month, current_date.day) in holidays_2026
        is_pre_holiday = (current_date.year, current_date.month, current_date.day) in pre_holidays_2026
        is_working_saturday = (current_date.year, current_date.month, current_date.day) in working_saturdays_2026
        
        if is_holiday:
            day_type = "holiday"
            day_name = "Праздник"
        elif is_pre_holiday:
            day_type = "pre_holiday"
            day_name = "Предпр"
        elif is_working_saturday:
            day_type = "work_saturday"
            day_name = "Раб.сб"
        elif weekday >= 5:  # Суббота или воскресенье
            day_type = "weekend"
            day_name = "Выходной"
        else:
            day_type = "workday"
            day_name = "Рабочий"
        
        calendar[date_key] = {
            'date': current_date,
            'day': current_date.day,
            'month': current_date.month,
            'year': current_date.year,
            'weekday': weekday,
            'day_type': day_type,
            'day_name': day_name,
            'is_working': day_type in ['workday', 'work_saturday', 'pre_holiday']
        }
    
    return calendar

def create_employees_sheet(ws, max_employees=MAX_EMPLOYEES):
    """Создает лист сотрудников с запасными строками"""
    
    # Стили
    header_fill = PatternFill(start_color="1F497D", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Заголовок компании
    ws.merge_cells('A1:K1')
    company_cell = ws['A1']
    company_cell.value = f"{COMPANY_NAME} - ГРАФИК ОТПУСКОВ {YEAR}"
    company_cell.font = Font(bold=True, size=14, color="1F497D")
    company_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Заголовки таблицы
    headers = [
        "№ п/п", "ФАМИЛИЯ ИМЯ ОТЧЕСТВО",
        "ОТПУСК 1", "ОТПУСК 1", "ОТПУСК 1",
        "ОТПУСК 2", "ОТПУСК 2", "ОТПУСК 2",
        "ОТПУСК 3", "ОТПУСК 3", "ОТПУСК 3"
    ]
    
    for col, header in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=header)
    
    # Объединяем ячейки для заголовков отпусков
    ws.merge_cells('C3:E3')
    ws.merge_cells('F3:H3')
    ws.merge_cells('I3:K3')
    
    # Подзаголовки
    sub_headers = ["", "",
                  "Начало", "Конец", "Дней",
                  "Начало", "Конец", "Дней",
                  "Начало", "Конец", "Дней"]
    
    for col, header in enumerate(sub_headers, 1):
        if header:
            ws.cell(row=4, column=col, value=header)
    
    # Применяем стили к заголовкам (строки 3 и 4)
    for row in [3, 4]:
        for col in range(1, 12):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
                cell.border = thin_border
    
    # Настраиваем ширину столбцов
    column_widths = [6, 35, 12, 12, 8, 12, 12, 8, 12, 12, 8]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    
    # Создаем строки для сотрудников (с запасом)
    start_row = 5  # Начало данных сотрудников
    
    for i in range(max_employees):
        row_num = start_row + i
        
        # Номер по порядку
        ws.cell(row=row_num, column=1, value=i+1)
        ws.cell(row=row_num, column=1).alignment = center_align
        
        # ФИО (оставляем пустым для будущего заполнения)
        ws.cell(row=row_num, column=2, value="")
        
        # Даты отпусков (оставляем пустыми)
        for col in [3, 4, 6, 7, 9, 10]:
            ws.cell(row=row_num, column=col, value="")
        
        # Поля для дней отпуска (будут заполняться макросом)
        for col in [5, 8, 11]:
            ws.cell(row=row_num, column=col, value="")
        
        # Применяем границы ко всем ячейкам строки
        for col in range(1, 12):
            cell = ws.cell(row=row_num, column=col)
            cell.border = thin_border
            if col >= 3:  # Все столбцы кроме № и ФИО
                cell.alignment = center_align
        
        # Закрашиваем строки через одну для удобства чтения
        if row_num % 2 == 0:
            row_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
            for col in range(1, 12):
                ws.cell(row=row_num, column=col).fill = row_fill
    
    # Добавляем строку для итогов
    total_row = start_row + max_employees + 1
    ws.cell(row=total_row, column=1, value="ИТОГО дней отпуска:")
    ws.cell(row=total_row, column=1).font = Font(bold=True)
    
    # Формула для подсчета дней (будет работать в Excel)
    ws.cell(row=total_row, column=5, value=f"=SUM(E{start_row}:E{start_row + max_employees - 1})")
    ws.cell(row=total_row, column=8, value=f"=SUM(H{start_row}:H{start_row + max_employees - 1})")
    ws.cell(row=total_row, column=11, value=f"=SUM(K{start_row}:K{start_row + max_employees - 1})")
    
    # Итоговая формула
    ws.cell(row=total_row, column=12, value="Общий итог:")
    ws.cell(row=total_row, column=13, value=f"=SUM(E{total_row},H{total_row},K{total_row})")
    ws.cell(row=total_row, column=13).font = Font(bold=True)
    
    return start_row

def create_schedule_sheet(ws, calendar, max_employees=MAX_EMPLOYEES):
    """Создает лист графика отпусков"""
    
    # Стили
    header_fill = PatternFill(start_color="1F497D", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Заголовок графика
    ws.merge_cells('A1:Z1')
    title_cell = ws['A1']
    title_cell.value = f"{COMPANY_NAME} - ГРАФИК ОТПУСКОВ НА {YEAR} ГОД"
    title_cell.font = Font(bold=True, size=14, color="1F497D")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Заголовки столбцов
    ws['A3'] = "№ п/п"
    ws['B3'] = "ФИО СОТРУДНИКА"
    
    for col in ['A', 'B']:
        cell = ws[f'{col}3']
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 35
    
    # Создаем строки для сотрудников (с запасом)
    start_row = 4  # Начало данных сотрудников в графике
    
    for i in range(max_employees):
        row_num = start_row + i
        
        # Номер
        ws.cell(row=row_num, column=1, value=i+1)
        ws.cell(row=row_num, column=1).alignment = center_align
        
        # ФИО (оставляем пустым)
        ws.cell(row=row_num, column=2, value="")
        ws.cell(row=row_num, column=2).alignment = Alignment(vertical="center")
        
        # Границы
        for col in [1, 2]:
            ws.cell(row=row_num, column=col).border = thin_border
        
        # Закрашивание через строку
        if row_num % 2 == 0:
            row_fill = PatternFill(start_color="F8F8F8", fill_type="solid")
            for col in [1, 2]:
                ws.cell(row=row_num, column=col).fill = row_fill
    
    # Создаем календарь на листе
    last_col = create_calendar_on_sheet(ws, calendar, start_row)
    
    # Добавляем кнопку обновления
    button_row = start_row + max_employees + 2
    ws.cell(row=button_row, column=1, value="🔄 ОБНОВИТЬ ГРАФИК ОТПУСКОВ")
    button_cell = ws.cell(row=button_row, column=1)
    button_cell.font = Font(bold=True, color="FFFFFF", size=12)
    button_cell.fill = PatternFill(start_color="4CAF50", fill_type="solid")
    button_cell.alignment = center_align
    button_cell.border = thin_border
    
    ws.merge_cells(f'A{button_row}:B{button_row}')
    
    # Инструкция
    instruction = "Внесите даты отпусков на листе 'СОТРУДНИКИ', затем нажмите Alt+F8 и запустите макрос 'ОбновитьГрафик'"
    ws.cell(row=button_row + 1, column=1, value=instruction)
    ws.cell(row=button_row + 1, column=1).font = Font(color="666666", italic=True)
    
    return last_col

def create_calendar_on_sheet(ws, calendar, schedule_start_row=4):
    """Создает календарь на листе графика"""
    
    # Цвета месяцев
    month_colors = {
        1: "4F81BD", 2: "8064A2", 3: "9BBB59", 4: "C0504D",
        5: "F79646", 6: "1F497D", 7: "948A54", 8: "31869B",
        9: "E26B0A", 10: "60497A", 11: "C00000", 12: "366092"
    }
    
    # Названия месяцев
    month_names = {
        1: "ЯНВ", 2: "ФЕВ", 3: "МАР", 4: "АПР",
        5: "МАЙ", 6: "ИЮН", 7: "ИЮЛ", 8: "АВГ",
        9: "СЕН", 10: "ОКТ", 11: "НОЯ", 12: "ДЕК"
    }
    
    # Дни недели
    weekday_names = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    
    # Группируем дни по месяцам
    months = {}
    for date_info in calendar.values():
        month = date_info['month']
        if month not in months:
            months[month] = []
        months[month].append(date_info)
    
    sorted_months = sorted(months.keys())
    current_col = 3  # Начинаем с колонки C
    
    # Заголовки для календаря (строка 2)
    for month_num in sorted_months:
        month_days = months[month_num]
        start_col = current_col
        end_col = current_col + len(month_days) - 1
        
        # Объединяем для названия месяца
        start_letter = get_column_letter(start_col)
        end_letter = get_column_letter(end_col)
        ws.merge_cells(f"{start_letter}2:{end_letter}2")
        
        # Название месяца
        month_cell = ws.cell(row=2, column=start_col)
        month_cell.value = month_names[month_num]
        month_cell.fill = PatternFill(start_color=month_colors[month_num], fill_type="solid")
        month_cell.font = Font(color="FFFFFF", bold=True, size=11)
        month_cell.alignment = Alignment(horizontal="center", vertical="center")
        month_cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Дни месяца
        for i, day_info in enumerate(month_days):
            col = current_col + i
            
            # Строка 3: ЧИСЛО ДНЯ (видимое)
            day_cell = ws.cell(row=3, column=col, value=day_info['day'])
            day_cell.alignment = Alignment(horizontal="center", vertical="center")
            day_cell.font = Font(bold=True, size=9)
            day_cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Определяем цвет фона для типа дня
            bg_color = "FFFFFF"
            if day_info['day_type'] == 'holiday':
                bg_color = "FF9999"
            elif day_info['day_type'] == 'pre_holiday':
                bg_color = "FFFF99"
            elif day_info['day_type'] == 'work_saturday':
                bg_color = "CCFFCC"
            elif day_info['day_type'] == 'weekend':
                bg_color = "E6E6E6"
            
            day_cell.fill = PatternFill(start_color=bg_color, fill_type="solid")
            
            # Строка 4: ДЕНЬ НЕДЕЛИ
            weekday = weekday_names[day_info['weekday']]
            
            # Добавляем символы для особых дней
            symbol = ""
            if day_info['day_type'] == 'holiday':
                symbol = " ✶"
            elif day_info['day_type'] == 'pre_holiday':
                symbol = " ◐"
            elif day_info['day_type'] == 'work_saturday':
                symbol = " ⚒"
            
            weekday_cell = ws.cell(row=4, column=col, value=f"{weekday}{symbol}")
            weekday_cell.alignment = Alignment(horizontal="center", vertical="center")
            weekday_cell.font = Font(size=9)
            weekday_cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            weekday_cell.fill = PatternFill(start_color=bg_color, fill_type="solid")
            
            # Скрытая строка 5: полная дата для макроса
            date_cell = ws.cell(row=5, column=col)
            date_cell.value = day_info['date']  # Полная дата
            date_cell.number_format = 'DD.MM.YYYY'
            date_cell.font = Font(size=1, color="FFFFFF")  # Почти невидимый
            
            # Ширина столбца
            ws.column_dimensions[get_column_letter(col)].width = 4.5
        
        current_col += len(month_days)
    
    # Скрываем строку 5 с датами
    ws.row_dimensions[5].hidden = True
    
    return current_col - 1

def create_instructions_sheet(ws):
    """Создает лист с инструкциями"""
    
    # Заголовок
    ws.merge_cells('A1:E1')
    title_cell = ws['A1']
    title_cell.value = f"ИНСТРУКЦИЯ ПО РАБОТЕ С ГРАФИКОМ ОТПУСКОВ {YEAR}"
    title_cell.font = Font(bold=True, size=14, color="1F497D")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    instructions = [
        ["РАЗДЕЛ 1: ОСНОВНЫЕ ШАГИ", "", "", "", ""],
        ["1. ЗАПОЛНЕНИЕ ДАННЫХ", "", "", "", ""],
        ["• Откройте лист 'СОТРУДНИКИ'", "", "", "", ""],
        ["• В столбце B введите ФИО сотрудников", "", "", "", ""],
        ["• В столбцах C, D, F, G, I, J введите даты отпусков", "", "", "", ""],
        ["• Формат дат: ДД.ММ.ГГГГ (например, 15.01.2026)", "", "", "", ""],
        ["• Можно оставлять строки пустыми для будущих сотрудников", "", "", "", ""],
        ["", "", "", "", ""],
        ["2. ОБНОВЛЕНИЕ ГРАФИКА", "", "", "", ""],
        ["• После заполнения дат перейдите на лист 'ГРАФИК'", "", "", "", ""],
        ["• Нажмите Alt+F8 (или Developer → Macros)", "", "", "", ""],
        ["• Выберите макрос 'ОбновитьГрафик'", "", "", "", ""],
        ["• Нажмите 'Выполнить'", "", "", "", ""],
        ["• График автоматически обновится", "", "", "", ""],
        ["", "", "", "", ""],
        ["РАЗДЕЛ 2: ДОБАВЛЕНИЕ СОТРУДНИКОВ", "", "", "", ""],
        ["• В листе 'СОТРУДНИКИ' уже подготовлено 20 строк", "", "", "", ""],
        ["• Просто введите ФИО в пустую строку", "", "", "", ""],
        ["• Затем введите даты отпусков", "", "", "", ""],
        ["• Запустите макрос обновления", "", "", "", ""],
        ["• Новый сотрудник появится в графике", "", "", "", ""],
        ["", "", "", "", ""],
        ["РАЗДЕЛ 3: ОБОЗНАЧЕНИЯ В ГРАФИКЕ", "", "", "", ""],
        ["• Белый фон - рабочий день", "", "", "", ""],
        ["• Серый фон - выходной день", "", "", "", ""],
        ["• Красный фон + ✶ - праздничный день", "", "", "", ""],
        ["• Желтый фон + ◐ - предпраздничный день", "", "", "", ""],
        ["• Зеленый фон + ⚒ - рабочая суббота", "", "", "", ""],
        ["• Светло-зеленый + 'О' - отпуск сотрудника", "", "", "", ""],
        ["", "", "", "", ""],
        ["РАЗДЕЛ 4: ВАЖНЫЕ ПРИМЕЧАНИЯ", "", "", "", ""],
        ["• Макрос игнорирует строки с пустым ФИО", "", "", "", ""],
        ["• Даты проверяются на корректность", "", "", "", ""],
        ["• При ошибке в датах появится сообщение", "", "", "", ""],
        ["• Для сохранения макроса сохраните файл как .xlsm", "", "", "", ""],
        ["• Регулярно сохраняйте резервные копии", "", "", "", ""],
        ["", "", "", "", ""],
        ["ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ", "", "", "", ""],
        [f"Версия файла: 2.0 (Профессиональная)", "", "", "", ""],
        [f"Дата создания: {datetime.now().strftime('%d.%m.%Y %H:%M')}", "", "", "", ""],
        [f"Максимальное количество сотрудников: {MAX_EMPLOYEES}", "", "", "", ""],
        [f"Год графика: {YEAR}", "", "", "", ""],
        [f"Компания: {COMPANY_NAME}", "", "", "", ""],
    ]
    
    for row_idx, row_data in enumerate(instructions, start=3):
        for col_idx, cell_value in enumerate(row_data[:5], start=1):
            if cell_value:
                cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                
                # Форматирование заголовков
                if "РАЗДЕЛ" in cell_value or "ТЕХНИЧЕСКАЯ" in cell_value:
                    cell.font = Font(bold=True, size=12, color="1F497D")
                elif cell_value.startswith(("1.", "2.", "3.", "4.")):
                    cell.font = Font(bold=True, size=11, color="C00000")
                elif cell_value.startswith("•"):
                    cell.font = Font(size=10)
                elif "Версия" in cell_value or "Дата" in cell_value:
                    cell.font = Font(italic=True, color="666666")
    
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = 5
    ws.column_dimensions['D'].width = 5
    ws.column_dimensions['E'].width = 5

def create_macro_file():
    """Создает файл с улучшенным макросом"""
    
    macro_code = f'''Option Explicit
' УЛУЧШЕННЫЙ МАКРОС ДЛЯ ГРАФИКА ОТПУСКОВ
' ИГНОРИРУЕТ ПУСТЫЕ СТРОКИ, РАБОТАЕТ С {MAX_EMPLOYEES} СОТРУДНИКАМИ

Public Sub ОбновитьГрафик()
    Dim wsСотрудники As Worksheet
    Dim wsГрафик As Worksheet
    Dim последняяСтрока As Long
    Dim i As Long, col As Long
    Dim датаНачало As Date
    Dim датаКонец As Date
    Dim текущаяДата As Date
    Dim найденныйСтолбец As Long
    Dim днейОтпуска As Long
    Dim обработаноСотрудников As Integer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ОшибкаОбработки
    
    Set wsСотрудники = ThisWorkbook.Worksheets("СОТРУДНИКИ")
    Set wsГрафик = ThisWorkbook.Worksheets("ГРАФИК")
    
    ' 1. ОЧИСТКА СТАРЫХ ДАННЫХ В ГРАФИКЕ
    Call ОчиститьСтарыйГрафик(wsГрафик)
    
    ' 2. ОБРАБОТКА СОТРУДНИКОВ (начиная со строки 5)
    обработаноСотрудников = 0
    
    For i = 5 To {5 + MAX_EMPLOYEES - 1} ' Обрабатываем все зарезервированные строки
        ' ПРОВЕРКА: если ФИО пустое - пропускаем сотрудника
        If Trim(wsСотрудники.Cells(i, 2).Value) = "" Then
            ' Очищаем поля дней для пустых строк
            wsСотрудники.Cells(i, 5).ClearContents
            wsСотрудники.Cells(i, 8).ClearContents
            wsСотрудники.Cells(i, 11).ClearContents
        Else
            ' ОБРАБАТЫВАЕМ СОТРУДНИКА С ДАННЫМИ
            обработаноСотрудников = обработаноСотрудников + 1
            
            ' Копируем ФИО в график (строка в графике = i-1)
            wsГрафик.Cells(i - 1, 2).Value = wsСотрудники.Cells(i, 2).Value
            
            ' Первый отпуск (столбцы 3-5)
            Call ОбработатьОтпуск(wsСотрудники, wsГрафик, i, 3, 4, 5, i - 1)
            
            ' Второй отпуск (столбцы 6-8)
            Call ОбработатьОтпуск(wsСотрудники, wsГрафик, i, 6, 7, 8, i - 1)
            
            ' Третий отпуск (столбцы 9-11)
            Call ОбработатьОтпуск(wsСотрудники, wsГрафик, i, 9, 10, 11, i - 1)
        End If
    Next i
    
    ' 3. ОБНОВЛЕНИЕ ИТОГОВ
    Call ОбновитьИтоги(wsСотрудники)
    
    ' 4. АВТОПОДБОР ШИРИНЫ СТОЛБЦОВ
    wsГрафик.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' 5. ИНФОРМАЦИОННОЕ СООБЩЕНИЕ
    Dim сообщение As String
    сообщение = "График отпусков успешно обновлен!" & vbCrLf & vbCrLf
    сообщение = сообщение & "Обработано сотрудников: " & обработаноСотрудников & vbCrLf
    сообщение = сообщение & "Пустых строк проигнорировано: " & (MAX_EMPLOYEES - обработаноСотрудников)
    
    MsgBox сообщение, vbInformation, "Обновление завершено"
    Exit Sub
    
ОшибкаОбработки:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Ошибка при обновлении графика:" & vbCrLf & Err.Description, vbCritical, "Ошибка"
End Sub

Private Sub ОчиститьСтарыйГрафик(ws As Worksheet)
    Dim последнийСтолбец As Long
    Dim последняяСтрока As Long
    Dim i As Long, j As Long
    
    последнийСтолбец = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    последняяСтрока = {4 + MAX_EMPLOYEES - 1} ' Все строки сотрудников
    
    If последнийСтолбец > 2 Then
        For i = 4 To последняяСтрока ' Строки с сотрудниками
            For j = 3 To последнийСтолбец
                With ws.Cells(i, j)
                    .ClearContents
                    .Interior.ColorIndex = xlNone
                    .Font.ColorIndex = xlAutomatic
                    .Font.Bold = False
                End With
            Next j
        Next i
    End If
End Sub

Private Sub ОбработатьОтпуск(wsДанные As Worksheet, wsГрафик As Worksheet, _
                            строкаДанных As Long, столбецНачало As Long, _
                            столбецКонец As Long, столбецДни As Long, _
                            строкаГрафика As Long)
    Dim датаНачало As Date
    Dim датаКонец As Date
    Dim текущаяДата As Date
    Dim номерСтолбца As Long
    Dim днейОтпуска As Long
    
    On Error Resume Next
    датаНачало = CDate(wsДанные.Cells(строкаДанных, столбецНачало).Value)
    датаКонец = CDate(wsДанные.Cells(строкаДанных, столбецКонец).Value)
    On Error GoTo 0
    
    ' ПРОВЕРКА ВАЛИДНОСТИ ДАТ
    If IsDate(датаНачало) And IsDate(датаКонец) Then
        If датаКонец >= датаНачало Then
            ' РАСЧЕТ КОЛИЧЕСТВА ДНЕЙ
            днейОтпуска = DateDiff("d", датаНачало, датаКонец) + 1
            wsДанные.Cells(строкаДанных, столбецДни).Value = днейОтпуска
            
            ' ЦВЕТ ДЛЯ ОТПУСКА (светло-зеленый)
            Dim цветОтпуска As Long
            цветОтпуска = RGB(144, 238, 144)
            
            ' ОТМЕТКА ОТПУСКА В ГРАФИКЕ
            текущаяДата = датаНачало
            Do While текущаяДата <= датаКонец
                номерСтолбца = НайтиСтолбецПоДате(wsГрафик, текущаяДата)
                
                If номерСтолбца > 0 Then
                    With wsГрафик.Cells(строкаГрафика, номерСтолбца)
                        .Value = "О"
                        .Interior.Color = цветОтпуска
                        .Font.Bold = True
                        .Font.Color = RGB(0, 100, 0)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                End If
                
                текущаяДата = DateAdd("d", 1, текущаяДата)
            Loop
        Else
            wsДанные.Cells(строкаДанных, столбецДни).Value = "ОШИБКА: Дата конца раньше начала"
        End If
    Else
        ' ЕСЛИ ДАТЫ НЕВАЛИДНЫ - ОЧИЩАЕМ ПОЛЕ
        wsДанные.Cells(строкаДанных, столбецДни).ClearContents
    End If
End Sub

Private Function НайтиСтолбецПоДате(ws As Worksheet, искомаяДата As Date) As Long
    Dim col As Long
    Dim последнийСтолбец As Long
    
    последнийСтолбец = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 3 To последнийСтолбец
        ' ИЩЕМ В СКРЫТОЙ СТРОКЕ 5 (там полные даты)
        If ws.Cells(5, col).Value <> "" Then
            If IsDate(ws.Cells(5, col).Value) Then
                Dim датаВЯчейке As Date
                датаВЯчейке = CDate(ws.Cells(5, col).Value)
                
                ' СРАВНИВАЕМ ДАТЫ
                If Year(датаВЯчейке) = Year(искомаяДата) And _
                   Month(датаВЯчейке) = Month(искомаяДата) And _
                   Day(датаВЯчейке) = Day(искомаяДата) Then
                    НайтиСтолбецПоДате = col
                    Exit Function
                End If
            End If
        End If
    Next col
    
    НайтиСтолбецПоДате = 0 ' Дата не найдена
End Function

Private Sub ОбновитьИтоги(ws As Worksheet)
    Dim строкаИтогов As Long
    Dim i As Long
    Dim всегоДней1 As Long, всегоДней2 As Long, всегоДней3 As Long
    
    строкаИтогов = {5 + MAX_EMPLOYEES + 1} ' Строка после всех сотрудников
    
    ' ОБНУЛЯЕМ СЧЕТЧИКИ
    всегоДней1 = 0
    всегоДней2 = 0
    всегоДней3 = 0
    
    ' СУММИРУЕМ ДНИ ТОЛЬКО ДЛЯ ЗАПОЛНЕННЫХ СТРОК
    For i = 5 To {5 + MAX_EMPLOYEES - 1}
        If Trim(ws.Cells(i, 2).Value) <> "" Then ' Только если есть ФИО
            If IsNumeric(ws.Cells(i, 5).Value) Then всегоДней1 = всегоДней1 + ws.Cells(i, 5).Value
            If IsNumeric(ws.Cells(i, 8).Value) Then всегоДней2 = всегоДней2 + ws.Cells(i, 8).Value
            If IsNumeric(ws.Cells(i, 11).Value) Then всегоДней3 = всегоДней3 + ws.Cells(i, 11).Value
        End If
    Next i
    
    ' ОБЩИЙ ИТОГ
    Dim общийИтог As Long
    общийИтог = всегоДней1 + всегоДней2 + всегоДней3
    
    ' ЗАПИСЫВАЕМ РЕЗУЛЬТАТЫ
    ws.Cells(строкаИтогов, 1).Value = "ИТОГО дней отпуска:"
    ws.Cells(строкаИтогов, 1).Font.Bold = True
    
    ws.Cells(строкаИтогов, 5).Value = всегоДней1
    ws.Cells(строкаИтогов, 8).Value = всегоДней2
    ws.Cells(строкаИтогов, 11).Value = всегоДней3
    
    ' ОБЩИЙ ИТОГ
    ws.Cells(строкаИтогов + 1, 1).Value = "ОБЩИЙ ИТОГ:"
    ws.Cells(строкаИтогов + 1, 1).Font.Bold = True
    
    ws.Cells(строкаИтогов + 1, 5).Value = общийИтог
    ws.Cells(строкаИтогов + 1, 5).Font.Bold = True
    ws.Cells(строкаИтогов + 1, 5).HorizontalAlignment = xlRight
End Sub

Public Sub ТестМакроса()
    MsgBox "Макрос готов к работе! Запустите 'ОбновитьГрафик'.", vbInformation, "Тест"
End Sub
'''
    
    # Сохраняем макрос
    macro_filename = "макрос_профессиональный.txt"
    with open(macro_filename, "w", encoding="utf-8") as f:
        f.write(macro_code)
    
    return macro_filename

def create_vacation_schedule_pro():
    """Создает профессиональный график отпусков"""
    
    print("=" * 70)
    print(f"ПРОФЕССИОНАЛЬНЫЙ ГЕНЕРАТОР ГРАФИКА ОТПУСКОВ {YEAR}")
    print(f"Максимальное количество сотрудников: {MAX_EMPLOYEES}")
    print(f"Компания: {COMPANY_NAME}")
    print("=" * 70)
    
    # 1. ГЕНЕРИРУЕМ КАЛЕНДАРЬ
    print("\n📅 Генерирую производственный календарь...")
    calendar = get_russian_calendar(YEAR)
    
    # 2. СОЗДАЕМ ИМЯ ФАЙЛА
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"отпуск_{COMPANY_NAME.replace(' ', '_')}_{YEAR}_{timestamp}.xlsx"
    
    print(f"\n📁 Создаю файл: {filename}")
    
    # 3. СОЗДАЕМ КНИГУ EXCEL
    wb = Workbook()
    
    # Удаляем дефолтный лист
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    
    # 4. СОЗДАЕМ ЛИСТ "СОТРУДНИКИ" (с запасом строк)
    print("👥 Создаю лист СОТРУДНИКИ (20 строк с запасом)...")
    ws_employees = wb.create_sheet(title="СОТРУДНИКИ")
    create_employees_sheet(ws_employees)
    
    # 5. СОЗДАЕМ ЛИСТ "ГРАФИК"
    print("📊 Создаю лист ГРАФИК...")
    ws_schedule = wb.create_sheet(title="ГРАФИК")
    last_col = create_schedule_sheet(ws_schedule, calendar)
    
    # 6. СОЗДАЕМ ЛИСТ "ИНСТРУКЦИЯ"
    print("📋 Создаю лист ИНСТРУКЦИЯ...")
    ws_instructions = wb.create_sheet(title="ИНСТРУКЦИЯ")
    create_instructions_sheet(ws_instructions)
    
    # 7. СОЗДАЕМ ЛИСТ "ЛЕГЕНДА" (цветовые обозначения)
    print("🎨 Создаю лист ЛЕГЕНДА...")
    ws_legend = wb.create_sheet(title="ЛЕГЕНДА")
    
    # Заголовок легенды
    ws_legend.merge_cells('A1:C1')
    legend_title = ws_legend['A1']
    legend_title.value = "ЛЕГЕНДА - ОБОЗНАЧЕНИЯ В ГРАФИКЕ"
    legend_title.font = Font(bold=True, size=14, color="1F497D")
    legend_title.alignment = Alignment(horizontal="center")
    
    # Данные легенды
    legend_data = [
        ["Цвет/Символ", "Обозначение", "Описание"],
        ["Белый фон", "Рабочий день", "Обычный рабочий день (понедельник-пятница)"],
        ["Серый фон", "Выходной день", "Суббота, воскресенье"],
        ["Красный фон + ✶", "Праздничный день", "Государственный праздник, нерабочий день"],
        ["Желтый фон + ◐", "Предпраздничный день", "Сокращенный рабочий день (на 1 час)"],
        ["Зеленый фон + ⚒", "Рабочая суббота", "Перенесенная рабочая суббота"],
        ["Светло-зеленый + 'О'", "Отпуск сотрудника", "Период ежегодного оплачиваемого отпуска"],
        ["", "", ""],
        ["ПРИМЕЧАНИЯ:", "", ""],
        ["• Максимальное количество сотрудников: 20", "", ""],
        ["• Пустые строки игнорируются при обновлении", "", ""],
        ["• Можно добавлять новых сотрудников в пустые строки", "", ""],
        ["• Формат дат: ДД.ММ.ГГГГ", "", ""],
        ["• После добавления данных запустите макрос обновления", "", ""],
    ]
    
    for row_idx, row_data in enumerate(legend_data, start=3):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws_legend.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 3 or "ПРИМЕЧАНИЯ:" in value:
                cell.font = Font(bold=True)
    
    ws_legend.column_dimensions['A'].width = 20
    ws_legend.column_dimensions['B'].width = 20
    ws_legend.column_dimensions['C'].width = 40
    
    # 8. СОХРАНЯЕМ EXCEL ФАЙЛ
    print(f"\n💾 Сохраняю файл: {filename}")
    wb.save(filename)
    
    # 9. СОЗДАЕМ ПРОФЕССИОНАЛЬНЫЙ МАКРОС
    print("⚙️ Создаю профессиональный макрос VBA...")
    macro_file = create_macro_file()
    
    # 10. ВЫВОД ИНФОРМАЦИИ
    print("\n" + "=" * 70)
    print("✅ ПРОФЕССИОНАЛЬНЫЙ ФАЙЛ УСПЕШНО СОЗДАН!")
    print("=" * 70)
    
    print(f"\n📁 СОЗДАННЫЕ ФАЙЛЫ:")
    print(f"   1. {filename} - Основной Excel файл")
    print(f"   2. {macro_file} - Профессиональный макрос VBA")
    
    print(f"\n📊 ХАРАКТЕРИСТИКИ ФАЙЛА:")
    print(f"   • Компания: {COMPANY_NAME}")
    print(f"   • Год: {YEAR}")
    print(f"   • Максимальное количество сотрудников: {MAX_EMPLOYEES}")
    print(f"   • Зарезервированных строк: {MAX_EMPLOYEES}")
    print(f"   • Листов в файле: 4 (СОТРУДНИКИ, ГРАФИК, ИНСТРУКЦИЯ, ЛЕГЕНДА)")
    print(f"   • Дней в календаре: {len(calendar)}")
    
    print(f"\n🎯 ОСОБЕННОСТИ ЭТОЙ ВЕРСИИ:")
    print(f"   ✅ Подготовлено {MAX_EMPLOYEES} строк для сотрудников")
    print(f"   ✅ Макрос игнорирует строки с пустыми ФИО")
    print(f"   ✅ Можно добавлять новых сотрудников в пустые строки")
    print(f"   ✅ Профессиональное оформление")
    print(f"   ✅ Подробные инструкции и легенда")
    print(f"   ✅ Автоматический расчет итогов")
    
    print(f"\n🚀 КАК ИСПОЛЬЗОВАТЬ:")
    print(f"   1. Откройте файл {filename}")
    print(f"   2. Прочитайте инструкцию на листе 'ИНСТРУКЦИЯ'")
    print(f"   3. Заполните данные на листе 'СОТРУДНИКИ'")
    print(f"   4. Добавьте макрос из файла {macro_file}")
    print(f"   5. Запустите макрос 'ОбновитьГрафик'")
    
    print(f"\n💡 СОВЕТЫ:")
    print(f"   • Для новых сотрудников используйте пустые строки")
    print(f"   • Сохраняйте файл как .xlsm после добавления макроса")
    print(f"   • Регулярно делайте резервные копии")
    
    return filename, macro_file

def main():
    try:
        excel_file, macro_file = create_vacation_schedule_pro()
        
        print("\n" + "=" * 70)
        print("🎯 СТРУКТУРА ПРОЕКТА ГОТОВА!")
        print("=" * 70)
        
        print(f"\n📋 ЛИСТЫ В ФАЙЛЕ:")
        print(f"   1. СОТРУДНИКИ - {MAX_EMPLOYEES} строк с запасом")
        print(f"   2. ГРАФИК - визуальное отображение отпусков")
        print(f"   3. ИНСТРУКЦИЯ - подробное руководство")
        print(f"   4. ЛЕГЕНДА - цветовые обозначения")
        
        print(f"\n⚙️ ДЛЯ РАЗРАБОТЧИКОВ:")
        print(f"   Чтобы изменить настройки, отредактируйте:")
        print(f"   • MAX_EMPLOYEES = {MAX_EMPLOYEES} (макс. сотрудников)")
        print(f"   • COMPANY_NAME = '{COMPANY_NAME}' (название компании)")
        print(f"   • YEAR = {YEAR} (год графика)")
        
        input("\nНажмите Enter для завершения...")
        
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()