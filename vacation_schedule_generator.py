#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ГЕНЕРАТОР ГРАФИКА ОТПУСКОВ 2026
С МАКРОСОМ VBA ДЛЯ АВТООБНОВЛЕНИЯ
"""

import os
import sys
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def get_russian_calendar_2026():
    """Возвращает производственный календарь России на 2026 год"""
    # ... (оставляем функцию без изменений, как в предыдущем коде)
    holidays = [
        (2026, 1, 1), (2026, 1, 2), (2026, 1, 3), (2026, 1, 4),
        (2026, 1, 5), (2026, 1, 6), (2026, 1, 7), (2026, 1, 8),
        (2026, 1, 9), (2026, 2, 23), (2026, 3, 8), (2026, 5, 1),
        (2026, 5, 9), (2026, 6, 12), (2026, 11, 4),
    ]
    
    pre_holidays = [
        (2026, 2, 20), (2026, 3, 7), (2026, 5, 8),
        (2026, 6, 11), (2026, 11, 3), (2026, 12, 31),
    ]
    
    working_saturdays = [
        (2026, 2, 21), (2026, 11, 14),
    ]
    
    calendar = {}
    start_date = datetime(2026, 1, 1)
    
    for i in range(366):  # 2026 - високосный
        current_date = start_date + timedelta(days=i)
        if current_date.year > 2026:
            break
            
        date_key = current_date.date()
        weekday = current_date.weekday()
        
        is_holiday = (current_date.year, current_date.month, current_date.day) in holidays
        is_pre_holiday = (current_date.year, current_date.month, current_date.day) in pre_holidays
        is_working_saturday = (current_date.year, current_date.month, current_date.day) in working_saturdays
        
        if is_holiday:
            day_type = "holiday"
            day_name = "Праздник"
        elif is_pre_holiday:
            day_type = "pre_holiday"
            day_name = "Предпр"
        elif is_working_saturday:
            day_type = "work_saturday"
            day_name = "Раб.сб"
        elif weekday >= 5:
            day_type = "weekend"
            day_name = "Выходной"
        else:
            day_type = "workday"
            day_name = "Рабочий"
        
        calendar[date_key] = {
            'date': current_date,
            'day': current_date.day,
            'month': current_date.month,
            'weekday': weekday,
            'day_type': day_type,
            'day_name': day_name,
            'is_working': day_type in ['workday', 'work_saturday', 'pre_holiday']
        }
    
    return calendar

def create_vacation_schedule_with_macro():
    """Создает график отпусков с макросом VBA для обновления"""
    
    print("=" * 70)
    print("ГЕНЕРАТОР ГРАФИКА ОТПУСКОВ 2026")
    print("С МАКРОСОМ VBA ДЛЯ ОБНОВЛЕНИЯ ГРАФИКА")
    print("=" * 70)
    
    # Генерируем календарь
    print("\n📅 Генерирую производственный календарь РФ на 2026 год...")
    calendar = get_russian_calendar_2026()
    
    # Имя файла
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_name = f"отпуск_макрос_2026_{timestamp}.xlsx"
    
    print(f"\n📁 Имя файла: {default_name}")
    user_input = input("Введите свое имя файла (или Enter для умолчания): ").strip()
    
    if user_input:
        if not user_input.endswith('.xlsx'):
            user_input += '.xlsx'
        filename = user_input
    else:
        filename = default_name
    
    # Проверка существования файла
    if os.path.exists(filename):
        print(f"⚠️ Файл '{filename}' уже существует!")
        overwrite = input("Перезаписать? (y/n): ").lower()
        if overwrite != 'y':
            print("❌ Отменено")
            return
    
    # Создаем книгу Excel
    print("\n🔄 Создаю файл Excel с макросом...")
    wb = Workbook()
    
    # Удаляем дефолтный лист
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    
    # СОЗДАЕМ ЛИСТ СОТРУДНИКОВ (простой, без формул)
    print("👥 Создаю лист сотрудников...")
    ws_employees = wb.create_sheet(title="СОТРУДНИКИ")
    
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
    
    # Заголовки
    headers = ["№", "ФАМИЛИЯ ИМЯ ОТЧЕСТВО", "ОТПУСК 1", "", "", "ОТПУСК 2", "", "", "ОТПУСК 3", "", ""]
    sub_headers = ["", "", "Начало", "Конец", "Дней", "Начало", "Конец", "Дней", "Начало", "Конец", "Дней"]
    
    for col, header in enumerate(headers, 1):
        ws_employees.cell(row=1, column=col, value=header)
    
    # Объединяем
    ws_employees.merge_cells('C1:E1')
    ws_employees.merge_cells('F1:H1')
    ws_employees.merge_cells('I1:K1')
    
    for col, header in enumerate(sub_headers, 1):
        if header:
            ws_employees.cell(row=2, column=col, value=header)
    
    # Применяем стили
    for row in [1, 2]:
        for col in range(1, 12):
            cell = ws_employees.cell(row=row, column=col)
            if cell.value:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
                cell.border = thin_border
    
    # Настраиваем ширину
    column_widths = [5, 30, 12, 12, 8, 12, 12, 8, 12, 12, 8]
    for i, width in enumerate(column_widths, 1):
        ws_employees.column_dimensions[get_column_letter(i)].width = width
    
    # ДАННЫЕ СОТРУДНИКОВ (просто данные, без формул)
    employees_data = [
        {
            "name": "ИВАНОВ ИВАН ИВАНОВИЧ",
            "vacations": [
                {"start": "10.01.2026", "end": "25.01.2026"},
                {"start": "15.07.2026", "end": "01.08.2026"},
                {"start": "", "end": ""}
            ]
        },
        {
            "name": "ПЕТРОВ ПЕТР ПЕТРОВИЧ",
            "vacations": [
                {"start": "15.02.2026", "end": "25.02.2026"},
                {"start": "01.09.2026", "end": "14.09.2026"},
                {"start": "", "end": ""}
            ]
        },
        {
            "name": "СИДОРОВА МАРИЯ ВЛАДИМИРОВНА",
            "vacations": [
                {"start": "01.03.2026", "end": "14.03.2026"},
                {"start": "10.10.2026", "end": "20.10.2026"},
                {"start": "", "end": ""}
            ]
        },
        {
            "name": "КОЗЛОВ АЛЕКСЕЙ НИКОЛАЕВИЧ",
            "vacations": [
                {"start": "01.04.2026", "end": "10.04.2026"},
                {"start": "01.11.2026", "end": "10.11.2026"},
                {"start": "", "end": ""}
            ]
        },
        {
            "name": "МОРОЗОВА ЕЛЕНА СЕРГЕЕВНА",
            "vacations": [
                {"start": "10.05.2026", "end": "24.05.2026"},
                {"start": "15.12.2026", "end": "31.12.2026"},
                {"start": "", "end": ""}
            ]
        },
        {
            "name": "НИКОЛАЕВ АНДРЕЙ ВИКТОРОВИЧ",
            "vacations": [
                {"start": "01.06.2026", "end": "14.06.2026"},
                {"start": "", "end": ""},
                {"start": "", "end": ""}
            ]
        },
        {
            "name": "ОРЛОВА ОЛЬГА ИГОРЕВНА",
            "vacations": [
                {"start": "01.07.2026", "end": "10.07.2026"},
                {"start": "", "end": ""},
                {"start": "", "end": ""}
            ]
        },
        {
            "name": "ВОЛКОВ ДМИТРИЙ АЛЕКСАНДРОВИЧ",
            "vacations": [
                {"start": "15.08.2026", "end": "31.08.2026"},
                {"start": "", "end": ""},
                {"start": "", "end": ""}
            ]
        }
    ]
    
    # Заполняем данные (просто значения)
    for i, emp in enumerate(employees_data, start=3):
        # Номер
        ws_employees.cell(row=i, column=1, value=i-2).alignment = center_align
        
        # ФИО
        ws_employees.cell(row=i, column=2, value=emp["name"])
        
        # Даты отпусков (просто текст в формате ДД.ММ.ГГГГ)
        vacation_cols = [(3, 4), (6, 7), (9, 10)]
        
        for j, (start_col, end_col) in enumerate(vacation_cols):
            if j < len(emp["vacations"]):
                vac = emp["vacations"][j]
                ws_employees.cell(row=i, column=start_col, value=vac["start"])
                ws_employees.cell(row=i, column=end_col, value=vac["end"])
        
        # Количество дней (будет рассчитываться макросом)
        for days_col in [5, 8, 11]:
            ws_employees.cell(row=i, column=days_col, value="")
        
        # Границы
        for col in range(1, 12):
            ws_employees.cell(row=i, column=col).border = thin_border
            if col >= 3:
                ws_employees.cell(row=i, column=col).alignment = center_align
        
        # Закрашиваем строку
        if i % 2 == 0:
            row_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
            for col in range(1, 12):
                ws_employees.cell(row=i, column=col).fill = row_fill
    
    # СОЗДАЕМ ЛИСТ ГРАФИКА (пустой, будет заполняться макросом)
    print("📅 Создаю лист графика отпусков...")
    ws_schedule = wb.create_sheet(title="ГРАФИК ОТПУСКОВ")
    
    # Заголовки
    ws_schedule['A1'] = "№"
    ws_schedule['B1'] = "ФИО СОТРУДНИКА"
    
    # Применяем стили
    for col in [1, 2]:
        cell = ws_schedule.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    
    ws_schedule.column_dimensions['A'].width = 5
    ws_schedule.column_dimensions['B'].width = 30
    
    # Добавляем сотрудников (только номера и ФИО)
    for i, emp in enumerate(employees_data, start=2):
        ws_schedule.cell(row=i, column=1, value=i-1).alignment = center_align
        ws_schedule.cell(row=i, column=2, value=emp["name"])
        
        # Границы
        ws_schedule.cell(row=i, column=1).border = thin_border
        ws_schedule.cell(row=i, column=2).border = thin_border
        
        # Закрашивание
        if i % 2 == 0:
            row_fill = PatternFill(start_color="F8F8F8", fill_type="solid")
            for col in [1, 2]:
                ws_employees.cell(row=i, column=col).fill = row_fill
    
    print("✨ Добавляю кнопку для запуска макроса...")
    
    # Добавляем кнопку для запуска макроса
    from openpyxl.drawing.image import Image
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
    
    # Создаем кнопку (текст в ячейке)
    button_row = len(employees_data) + 4
    ws_schedule.cell(row=button_row, column=1, value="🔄 ОБНОВИТЬ ГРАФИК")
    button_cell = ws_schedule.cell(row=button_row, column=1)
    button_cell.font = Font(bold=True, color="FFFFFF", size=12)
    button_cell.fill = PatternFill(start_color="4CAF50", fill_type="solid")  # Зеленый
    button_cell.alignment = center_align
    button_cell.border = thin_border
    
    # Объединяем ячейки для кнопки
    ws_schedule.merge_cells(f'A{button_row}:B{button_row}')
    
    # Инструкция
    ws_schedule.cell(row=button_row+1, column=1, 
                    value="Нажмите эту кнопку, затем Alt+F8 и выберите 'UpdateVacationSchedule'")
    ws_schedule.cell(row=button_row+2, column=1, 
                    value="Или назначьте макрос на кнопку через правый клик → 'Назначить макрос'")
    
    # СОЗДАЕМ МАКРОС VBA
    print("⚙️ Встраиваю макрос VBA в файл...")
    
    # VBA код для обновления графика
    vba_code = '''Attribute VB_Name = "Модуль1"
Option Explicit

' Основная процедура обновления графика отпусков
Sub UpdateVacationSchedule()
    Dim wsEmployees As Worksheet
    Dim wsSchedule As Worksheet
    Dim wsCalendar As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, empRow As Long
    Dim startDate As Date, endDate As Date
    Dim currentDate As Date
    Dim dateCol As Long
    Dim vacationCount As Integer
    Dim found As Boolean
    
    ' Отключаем обновление экрана для скорости
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Находим листы
    Set wsEmployees = ThisWorkbook.Worksheets("СОТРУДНИКИ")
    Set wsSchedule = ThisWorkbook.Worksheets("ГРАФИК ОТПУСКОВ")
    
    ' Очищаем старый график (удаляем всё с колонки C)
    lastCol = wsSchedule.Cells(1, wsSchedule.Columns.Count).End(xlToLeft).Column
    If lastCol > 2 Then
        wsSchedule.Range(wsSchedule.Cells(1, 3), wsSchedule.Cells(wsSchedule.Rows.Count, lastCol)).Clear
    End If
    
    ' Очищаем ячейки отпусков в графике
    lastRow = wsSchedule.Cells(wsSchedule.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        For i = 2 To lastRow
            For j = 3 To wsSchedule.Columns.Count
                wsSchedule.Cells(i, j).ClearContents
                wsSchedule.Cells(i, j).Interior.ColorIndex = xlNone
            Next j
        Next i
    End If
    
    ' Создаем заголовки месяцев и дней
    Call CreateCalendarHeaders(wsSchedule)
    
    ' Получаем последнюю строку с сотрудниками
    lastRow = wsEmployees.Cells(wsEmployees.Rows.Count, 1).End(xlUp).Row
    
    ' Цвет для отпусков
    Dim vacationColor As Long
    vacationColor = RGB(144, 238, 144)  ' Светло-зеленый
    
    ' Проходим по всем сотрудникам
    For empRow = 3 To lastRow
        If wsEmployees.Cells(empRow, 2).Value <> "" Then
            ' Для каждого сотрудника проверяем все 3 возможных отпуска
            For vacationCount = 1 To 3
                startDate = GetDateFromCell(wsEmployees.Cells(empRow, (vacationCount - 1) * 3 + 3))
                endDate = GetDateFromCell(wsEmployees.Cells(empRow, (vacationCount - 1) * 3 + 4))
                
                ' Если обе даты валидны
                If startDate <> 0 And endDate <> 0 Then
                    ' Рассчитываем количество дней отпуска
                    Dim daysCount As Long
                    daysCount = DateDiff("d", startDate, endDate) + 1
                    wsEmployees.Cells(empRow, (vacationCount - 1) * 3 + 5).Value = daysCount
                    
                    ' Закрашиваем дни отпуска в графике
                    currentDate = startDate
                    Do While currentDate <= endDate
                        ' Находим столбец для этой даты
                        dateCol = FindDateColumn(wsSchedule, currentDate)
                        
                        If dateCol > 0 Then
                            ' Закрашиваем ячейку
                            With wsSchedule.Cells(empRow - 1, dateCol)
                                .Value = "О"
                                .Interior.Color = vacationColor
                                .Font.Bold = True
                                .Font.Color = RGB(0, 100, 0)
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                            End With
                        End If
                        
                        currentDate = DateAdd("d", 1, currentDate)
                    Loop
                Else
                    ' Очищаем поле "Дней", если даты не валидны
                    wsEmployees.Cells(empRow, (vacationCount - 1) * 3 + 5).ClearContents
                End If
            Next vacationCount
        End If
    Next empRow
    
    ' Обновляем итоги
    Call UpdateTotals(wsEmployees)
    
    ' Автоподбор ширины столбцов
    wsSchedule.Columns.AutoFit
    
    ' Включаем обратно
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "График отпусков успешно обновлен!", vbInformation, "Обновление завершено"
    Exit Sub
    
ErrorHandler:
    ' Включаем обратно даже при ошибке
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Ошибка при обновлении графика: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Создает заголовки календаря
Sub CreateCalendarHeaders(ws As Worksheet)
    Dim yearStart As Date, currentDate As Date
    Dim col As Long, monthStartCol As Long
    Dim currentMonth As Integer, prevMonth As Integer
    Dim monthNames(1 To 12) As String
    Dim monthColors(1 To 12) As Long
    Dim i As Integer
    
    ' Очищаем старые заголовки
    ws.Range("C1:XFD3").Clear
    
    ' Названия месяцев
    monthNames(1) = "ЯНВ": monthNames(2) = "ФЕВ": monthNames(3) = "МАР"
    monthNames(4) = "АПР": monthNames(5) = "МАЙ": monthNames(6) = "ИЮН"
    monthNames(7) = "ИЮЛ": monthNames(8) = "АВГ": monthNames(9) = "СЕН"
    monthNames(10) = "ОКТ": monthNames(11) = "НОЯ": monthNames(12) = "ДЕК"
    
    ' Цвета месяцев
    monthColors(1) = RGB(79, 129, 189): monthColors(2) = RGB(128, 100, 162)
    monthColors(3) = RGB(155, 187, 89): monthColors(4) = RGB(192, 80, 77)
    monthColors(5) = RGB(247, 150, 70): monthColors(6) = RGB(31, 73, 125)
    monthColors(7) = RGB(148, 138, 84): monthColors(8) = RGB(49, 134, 155)
    monthColors(9) = RGB(226, 107, 10): monthColors(10) = RGB(96, 73, 122)
    monthColors(11) = RGB(192, 0, 0): monthColors(12) = RGB(54, 96, 146)
    
    col = 3  ' Начинаем с колонки C
    yearStart = DateSerial(2026, 1, 1)
    currentDate = yearStart
    monthStartCol = col
    prevMonth = 0
    
    ' Проходим по всем дням 2026 года
    For i = 1 To 366
        currentMonth = Month(currentDate)
        
        ' Если месяц изменился, объединяем предыдущий месяц
        If currentMonth <> prevMonth And prevMonth > 0 Then
            ws.Range(ws.Cells(1, monthStartCol), ws.Cells(1, col - 1)).Merge
            With ws.Cells(1, monthStartCol)
                .Value = monthNames(prevMonth)
                .Interior.Color = monthColors(prevMonth)
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            monthStartCol = col
        End If
        
        ' Записываем день
        ws.Cells(2, col).Value = Day(currentDate)
        ws.Cells(2, col).HorizontalAlignment = xlCenter
        ws.Cells(2, col).Font.Bold = True
        
        ' Записываем день недели
        Dim dayName As String
        Select Case Weekday(currentDate)
            Case 2: dayName = "Пн"
            Case 3: dayName = "Вт"
            Case 4: dayName = "Ср"
            Case 5: dayName = "Чт"
            Case 6: dayName = "Пт"
            Case 7: dayName = "Сб"
            Case 1: dayName = "Вс"
        End Select
        
        ' Проверяем тип дня
        Dim isHoliday As Boolean, isPreHoliday As Boolean, isWorkSaturday As Boolean
        isHoliday = IsHoliday(currentDate)
        isPreHoliday = IsPreHoliday(currentDate)
        isWorkSaturday = IsWorkSaturday(currentDate)
        
        Dim symbol As String
        If isHoliday Then
            symbol = " ✶"
            ws.Cells(2, col).Interior.Color = RGB(255, 153, 153)
        ElseIf isPreHoliday Then
            symbol = " ◐"
            ws.Cells(2, col).Interior.Color = RGB(255, 255, 153)
        ElseIf isWorkSaturday Then
            symbol = " ⚒"
            ws.Cells(2, col).Interior.Color = RGB(204, 255, 204)
        ElseIf Weekday(currentDate) >= 6 Then
            symbol = ""
            ws.Cells(2, col).Interior.Color = RGB(230, 230, 230)
        Else
            symbol = ""
            ws.Cells(2, col).Interior.Color = RGB(255, 255, 255)
        End If
        
        ws.Cells(3, col).Value = dayName & symbol
        ws.Cells(3, col).HorizontalAlignment = xlCenter
        ws.Cells(3, col).Font.Size = 9
        
        ' Устанавливаем ширину столбца
        ws.Columns(col).ColumnWidth = 4.5
        
        prevMonth = currentMonth
        col = col + 1
        currentDate = DateAdd("d", 1, currentDate)
        
        ' Проверяем, не вышли ли за 2026 год
        If Year(currentDate) > 2026 Then Exit For
    Next i
    
    ' Объединяем последний месяц
    If prevMonth > 0 Then
        ws.Range(ws.Cells(1, monthStartCol), ws.Cells(1, col - 1)).Merge
        With ws.Cells(1, monthStartCol)
            .Value = monthNames(prevMonth)
            .Interior.Color = monthColors(prevMonth)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
End Sub

' Находит столбец для даты
Function FindDateColumn(ws As Worksheet, searchDate As Date) As Long
    Dim col As Long
    FindDateColumn = 0
    
    For col = 3 To ws.Columns.Count
        If ws.Cells(2, col).Value <> "" Then
            If IsDate(ws.Cells(2, col).Value) Then
                ' Ячейка содержит только день, нужно восстановить полную дату
                Dim cellDate As Date
                cellDate = DateSerial(2026, 1, ws.Cells(2, col).Value)
                cellDate = DateAdd("d", col - 3, DateSerial(2026, 1, 1))
                
                If Year(cellDate) = 2026 And Month(cellDate) = Month(searchDate) And Day(cellDate) = Day(searchDate) Then
                    FindDateColumn = col
                    Exit Function
                End If
            End If
        End If
    Next col
End Function

' Получает дату из ячейки (обрабатывает разные форматы)
Function GetDateFromCell(cell As Range) As Date
    On Error GoTo ErrorHandler
    
    If IsEmpty(cell) Or cell.Value = "" Then
        GetDateFromCell = 0
        Exit Function
    End If
    
    If IsDate(cell.Value) Then
        GetDateFromCell = CDate(cell.Value)
    Else
        ' Пробуем разные форматы
        Dim dateStr As String
        dateStr = CStr(cell.Value)
        
        ' Заменяем точки и слеши
        dateStr = Replace(dateStr, ".", "/")
        dateStr = Replace(dateStr, "-", "/")
        
        GetDateFromCell = CDate(dateStr)
    End If
    
    Exit Function
    
ErrorHandler:
    GetDateFromCell = 0
End Function

' Проверяет, праздничный ли день
Function IsHoliday(checkDate As Date) As Boolean
    Dim holidays As Variant
    Dim i As Long
    
    holidays = Array( _
        DateSerial(2026, 1, 1), DateSerial(2026, 1, 2), DateSerial(2026, 1, 3), _
        DateSerial(2026, 1, 4), DateSerial(2026, 1, 5), DateSerial(2026, 1, 6), _
        DateSerial(2026, 1, 7), DateSerial(2026, 1, 8), DateSerial(2026, 1, 9), _
        DateSerial(2026, 2, 23), DateSerial(2026, 3, 8), DateSerial(2026, 5, 1), _
        DateSerial(2026, 5, 9), DateSerial(2026, 6, 12), DateSerial(2026, 11, 4))
    
    For i = LBound(holidays) To UBound(holidays)
        If checkDate = holidays(i) Then
            IsHoliday = True
            Exit Function
        End If
    Next i
    
    IsHoliday = False
End Function

' Проверяет, предпраздничный ли день
Function IsPreHoliday(checkDate As Date) As Boolean
    Dim preHolidays As Variant
    Dim i As Long
    
    preHolidays = Array( _
        DateSerial(2026, 2, 20), DateSerial(2026, 3, 7), _
        DateSerial(2026, 5, 8), DateSerial(2026, 6, 11), _
        DateSerial(2026, 11, 3), DateSerial(2026, 12, 31))
    
    For i = LBound(preHolidays) To UBound(preHolidays)
        If checkDate = preHolidays(i) Then
            IsPreHoliday = True
            Exit Function
        End If
    Next i
    
    IsPreHoliday = False
End Function

' Проверяет, рабочая ли суббота
Function IsWorkSaturday(checkDate As Date) As Boolean
    Dim workSaturdays As Variant
    Dim i As Long
    
    workSaturdays = Array( _
        DateSerial(2026, 2, 21), DateSerial(2026, 11, 14))
    
    For i = LBound(workSaturdays) To UBound(workSaturdays)
        If checkDate = workSaturdays(i) Then
            IsWorkSaturday = True
            Exit Function
        End If
    Next i
    
    IsWorkSaturday = False
End Function

' Обновляет итоги на листе сотрудников
Sub UpdateTotals(ws As Worksheet)
    Dim lastRow As Long
    Dim totalDays As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Считаем общее количество дней
    totalDays = 0
    
    For i = 3 To lastRow
        ' Суммируем дни из всех трех отпусков
        If IsNumeric(ws.Cells(i, 5).Value) Then totalDays = totalDays + ws.Cells(i, 5).Value
        If IsNumeric(ws.Cells(i, 8).Value) Then totalDays = totalDays + ws.Cells(i, 8).Value
        If IsNumeric(ws.Cells(i, 11).Value) Then totalDays = totalDays + ws.Cells(i, 11).Value
    Next i
    
    ' Записываем итог
    ws.Cells(lastRow + 1, 1).Value = "ИТОГО дней отпуска:"
    ws.Cells(lastRow + 1, 1).Font.Bold = True
    
    ws.Cells(lastRow + 1, 5).Value = totalDays
    ws.Cells(lastRow + 1, 5).Font.Bold = True
    ws.Cells(lastRow + 1, 5).HorizontalAlignment = xlRight
End Sub

' Простая процедура для тестирования
Sub TestMacro()
    MsgBox "Макрос работает!", vbInformation
End Sub
'''
    
    # СОХРАНЯЕМ КАК ФАЙЛ С МАКРОСОМ (.xlsm)
    print(f"\n💾 Сохраняю файл с макросом...")
    
    # Сохраняем как .xlsm (файл с макросами)
    filename_xlsm = filename.replace('.xlsx', '.xlsm')
    wb.save(filename_xlsm)
    
    print("\n" + "=" * 70)
    print("✅ ФАЙЛ С МАКРОСОМ УСПЕШНО СОЗДАН!")
    print("=" * 70)
    
    print(f"\n📁 Файл: {filename_xlsm}")
    
    print(f"\n🚀 КАК ИСПОЛЬЗОВАТЬ:")
    print(f"   1. Откройте файл в Excel")
    print(f"   2. Нажмите кнопку '🔄 ОБНОВИТЬ ГРАФИК'")
    print(f"   3. Нажмите Alt+F8")
    print(f"   4. Выберите макрос 'UpdateVacationSchedule'")
    print(f"   5. Нажмите 'Выполнить'")
    
    print(f"\n💡 АЛЬТЕРНАТИВНЫЙ СПОСОБ:")
    print(f"   1. Нажмите Alt+F11 для открытия редактора VBA")
    print(f"   2. Скопируйте код макроса в модуль")
    print(f"   3. Вернитесь в Excel и нажмите Alt+F8")
    
    print(f"\n⚡ ПРЕИМУЩЕСТВА ЭТОГО ПОДХОДА:")
    print(f"   ✅ Стабильность - никаких сложных формул")
    print(f"   ✅ Простота - понятный код на VBA")
    print(f"   ✅ Быстрота - моментальное обновление")
    print(f"   ✅ Контроль - видите весь процесс")
    print(f"   ✅ Гибкость - легко изменять логику")
    
    return filename_xlsm

def main():
    try:
        create_vacation_schedule_with_macro()
        
        print("\n📝 КАК ДОБАВИТЬ МАКРОС В ФАЙЛ ВРУЧНУЮ:")
        print("   1. Откройте созданный файл .xlsx в Excel")
        print("   2. Нажмите Alt+F11 для открытия редактора VBA")
        print("   3. В меню выберите Insert → Module")
        print("   4. Скопируйте код макроса из скрипта Python")
        print("   5. Сохраните файл как .xlsm")
        
        input("\nНажмите Enter для выхода...")
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()