#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ПОЛНЫЙ ИСПРАВЛЕННЫЙ ГЕНЕРАТОР ГРАФИКА ОТПУСКОВ
БЕЗ СМЕЩЕНИЙ ДАТ, С КОРРЕКТНЫМ МАКРОСОМ
"""

import os
import sys
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def get_russian_calendar_2026():
    """Возвращает производственный календарь России на 2026 год"""
    
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
    
    for i in range(366):
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
            'year': current_date.year,
            'weekday': weekday,
            'day_type': day_type,
            'day_name': day_name,
            'is_working': day_type in ['workday', 'work_saturday', 'pre_holiday']
        }
    
    return calendar

def create_calendar_headers(ws, calendar):
    """Создает заголовки календаря на листе (ИСПРАВЛЕНО - без смещений)"""
    
    # Цвета месяцев
    month_colors = {
        1: "4F81BD", 2: "8064A2", 3: "9BBB59", 4: "C0504D",
        5: "F79646", 6: "1F497D", 7: "948A54", 8: "31869B",
        9: "E26B0A", 10: "60497A", 11: "C00000", 12: "366092"
    }
    
    # Группируем дни по месяцам
    months = {}
    for date_info in calendar.values():
        month = date_info['month']
        if month not in months:
            months[month] = []
        months[month].append(date_info)
    
    sorted_months = sorted(months.keys())
    current_col = 3  # Начинаем с колонки C
    
    month_names = {
        1: "ЯНВ", 2: "ФЕВ", 3: "МАР", 4: "АПР",
        5: "МАЙ", 6: "ИЮН", 7: "ИЮЛ", 8: "АВГ",
        9: "СЕН", 10: "ОКТ", 11: "НОЯ", 12: "ДЕК"
    }
    
    weekday_names = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    
    # Создаем структуру для хранения дат по столбцам (для макроса)
    date_column_map = {}
    
    for month_num in sorted_months:
        month_days = months[month_num]
        start_col = current_col
        end_col = current_col + len(month_days) - 1
        
        # Объединяем для названия месяца (строка 1)
        start_letter = get_column_letter(start_col)
        end_letter = get_column_letter(end_col)
        ws.merge_cells(f"{start_letter}1:{end_letter}1")
        
        # Название месяца в строке 1
        month_cell = ws[f"{start_letter}1"]
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
            date_obj = day_info['date']
            
            # Сохраняем соответствие дата -> столбец (для макроса)
            date_column_map[date_obj.date()] = col
            
            # СТРОКА 2: ЧИСЛО ДНЯ (видимое)
            day_cell = ws.cell(row=2, column=col, value=day_info['day'])
            day_cell.alignment = Alignment(horizontal="center", vertical="center")
            day_cell.font = Font(bold=True, size=9)
            day_cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # СТРОКА 3: ДЕНЬ НЕДЕЛИ
            weekday = weekday_names[day_info['weekday']]
            
            # Добавляем символы для особых дней
            symbol = ""
            bg_color = "FFFFFF"
            text_color = "000000"
            font_style = Font(size=9, color=text_color)
            
            if day_info['day_type'] == 'holiday':
                symbol = " ✶"
                bg_color = "FF9999"
                font_style = Font(size=9, color="000000", bold=True)
            elif day_info['day_type'] == 'pre_holiday':
                symbol = " ◐"
                bg_color = "FFFF99"
                font_style = Font(size=9, color="000000", italic=True)
            elif day_info['day_type'] == 'work_saturday':
                symbol = " ⚒"
                bg_color = "CCFFCC"
                font_style = Font(size=9, color="006600", bold=True)
            elif day_info['day_type'] == 'weekend':
                bg_color = "E6E6E6"
            
            weekday_cell = ws.cell(row=3, column=col, value=f"{weekday}{symbol}")
            weekday_cell.alignment = Alignment(horizontal="center", vertical="center")
            weekday_cell.font = font_style
            weekday_cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            weekday_cell.fill = PatternFill(start_color=bg_color, fill_type="solid")
            
            # Ширина столбца
            ws.column_dimensions[get_column_letter(col)].width = 4.5
            
            # Скрытая строка 4: полная дата для макроса (скрыта)
            date_cell = ws.cell(row=4, column=col)
            date_cell.value = date_obj  # Полная дата
            date_cell.number_format = 'DD.MM.YYYY'  # Формат даты
            date_cell.font = Font(size=1, color="FFFFFF")  # Почти невидимый
        
        current_col += len(month_days)
    
    # Скрываем строку 4 с датами
    ws.row_dimensions[4].hidden = True
    
    return current_col - 1, date_column_map

def create_vacation_schedule():
    """Создает полный график отпусков без смещений"""
    
    print("=" * 70)
    print("ГЕНЕРАТОР ГРАФИКА ОТПУСКОВ 2026 (ИСПРАВЛЕННЫЙ)")
    print("=" * 70)
    
    # 1. ГЕНЕРИРУЕМ КАЛЕНДАРЬ
    print("\n📅 Генерирую производственный календарь РФ на 2026 год...")
    calendar = get_russian_calendar_2026()
    
    # 2. ИМЯ ФАЙЛА
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"отпуск_исправленный_{timestamp}.xlsx"
    
    print(f"\n📁 Создаю файл: {filename}")
    
    # 3. СОЗДАЕМ КНИГУ EXCEL
    wb = Workbook()
    
    # Удаляем дефолтный лист
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    
    # 4. СОЗДАЕМ ЛИСТ "СОТРУДНИКИ"
    print("👥 Создаю лист СОТРУДНИКИ...")
    ws_data = wb.create_sheet(title="СОТРУДНИКИ")
    
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
    headers = ["№", "ФИО", "Отпуск1 Начало", "Отпуск1 Конец", "Дней",
               "Отпуск2 Начало", "Отпуск2 Конец", "Дней",
               "Отпуск3 Начало", "Отпуск3 Конец", "Дней"]
    
    for col, header in enumerate(headers, 1):
        cell = ws_data.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    
    # Ширина столбцов
    widths = [5, 30, 12, 12, 8, 12, 12, 8, 12, 12, 8]
    for i, width in enumerate(widths, 1):
        ws_data.column_dimensions[get_column_letter(i)].width = width
    
    # ТЕСТОВЫЕ ДАННЫЕ
    employees = [
        [1, "ИВАНОВ ИВАН ИВАНОВИЧ", "10.01.2026", "25.01.2026", "",
         "15.07.2026", "01.08.2026", "", "", "", ""],
        [2, "ПЕТРОВ ПЕТР ПЕТРОВИЧ", "15.02.2026", "25.02.2026", "",
         "01.09.2026", "14.09.2026", "", "", "", ""],
        [3, "СИДОРОВА МАРИЯ ВЛАДИМИРОВНА", "01.03.2026", "14.03.2026", "",
         "10.10.2026", "20.10.2026", "", "", "", ""],
        [4, "КОЗЛОВ АЛЕКСЕЙ НИКОЛАЕВИЧ", "01.04.2026", "10.04.2026", "",
         "01.11.2026", "10.11.2026", "", "", "", ""],
    ]
    
    for row_idx, emp in enumerate(employees, start=2):
        for col_idx, value in enumerate(emp, start=1):
            cell = ws_data.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(
                horizontal="center" if col_idx != 2 else "left",
                vertical="center"
            )
            cell.border = thin_border
        
        if row_idx % 2 == 0:
            for col in range(1, 12):
                ws_data.cell(row=row_idx, column=col).fill = PatternFill(
                    start_color="F2F2F2", fill_type="solid"
                )
    
    # 5. СОЗДАЕМ ЛИСТ "ГРАФИК"
    print("📊 Создаю лист ГРАФИК...")
    ws_graph = wb.create_sheet(title="ГРАФИК")
    
    # Заголовки графика
    ws_graph['A1'] = "№"
    ws_graph['B1'] = "ФИО"
    
    for col in ['A', 'B']:
        cell = ws_graph[f'{col}1']
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    
    ws_graph.column_dimensions['A'].width = 5
    ws_graph.column_dimensions['B'].width = 30
    
    # Добавляем сотрудников (начиная со строки 5)
    for i, emp in enumerate(employees, start=1):
        ws_graph.cell(row=i+4, column=1, value=emp[0])  # Строка 5 для первого сотрудника
        ws_graph.cell(row=i+4, column=1).alignment = center_align
        
        ws_graph.cell(row=i+4, column=2, value=emp[1])
        ws_graph.cell(row=i+4, column=2).alignment = Alignment(vertical="center")
        
        for col in [1, 2]:
            ws_graph.cell(row=i+4, column=col).border = thin_border
        
        if (i+4) % 2 == 0:
            for col in [1, 2]:
                ws_graph.cell(row=i+4, column=col).fill = PatternFill(
                    start_color="F8F8F8", fill_type="solid"
                )
    
    # 6. СОЗДАЕМ КАЛЕНДАРЬ (исправленный, без смещений)
    print("📅 Создаю календарь (январь начинается с колонки C)...")
    last_col, date_map = create_calendar_headers(ws_graph, calendar)
    
    # 7. СОЗДАЕМ КНОПКУ ДЛЯ МАКРОСА
    print("🔄 Добавляю кнопку для макроса...")
    
    button_row = len(employees) + 6
    ws_graph.cell(row=button_row, column=1, value="🔄 ОБНОВИТЬ ГРАФИК")
    button_cell = ws_graph.cell(row=button_row, column=1)
    button_cell.font = Font(bold=True, color="FFFFFF", size=12)
    button_cell.fill = PatternFill(start_color="4CAF50", fill_type="solid")
    button_cell.alignment = center_align
    button_cell.border = thin_border
    
    ws_graph.merge_cells(f'A{button_row}:B{button_row}')
    
    # Инструкция
    ws_graph.cell(row=button_row+1, column=1, 
                 value="Нажмите Alt+F8 и выберите 'ОбновитьГрафик'")
    
    # 8. СОХРАНЯЕМ ФАЙЛ
    print(f"\n💾 Сохраняю файл: {filename}")
    wb.save(filename)
    
    # 9. СОЗДАЕМ МАКРОС
    print("⚙️ Создаю макрос VBA...")
    
    # МАКРОС БЕЗ ПРОБЛЕМНОЙ СТРОКИ ATTRIBUTE
    macro_code = '''Option Explicit
' МАКРОС ДЛЯ ОБНОВЛЕНИЯ ГРАФИКА ОТПУСКОВ
' РАБОТАЕТ С ИСПРАВЛЕННОЙ СТРУКТУРОЙ ФАЙЛА

Public Sub ОбновитьГрафик()
    Dim wsData As Worksheet
    Dim wsGraph As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim targetCol As Long
    Dim daysCount As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Set wsData = ThisWorkbook.Worksheets("СОТРУДНИКИ")
    Set wsGraph = ThisWorkbook.Worksheets("ГРАФИК")
    
    ' 1. Очищаем старые отпуска (начиная со строки 5, столбцы C и дальше)
    Call ОчиститьОтпуска(wsGraph)
    
    ' 2. Находим последнего сотрудника
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' 3. Цвет для отпусков
    Dim vacationColor As Long
    vacationColor = RGB(144, 238, 144) ' Светло-зеленый
    
    ' 4. Обрабатываем каждого сотрудника
    For i = 2 To lastRow
        If wsData.Cells(i, 2).Value <> "" Then
            ' Первый отпуск (столбцы 3-5)
            Call ОбработатьОтпускСотрудника(wsData, wsGraph, i, 3, 4, 5, i + 3, vacationColor)
            
            ' Второй отпуск (столбцы 6-8)
            Call ОбработатьОтпускСотрудника(wsData, wsGraph, i, 6, 7, 8, i + 3, vacationColor)
            
            ' Третий отпуск (столбцы 9-11)
            Call ОбработатьОтпускСотрудника(wsData, wsGraph, i, 9, 10, 11, i + 3, vacationColor)
        End If
    Next i
    
    ' 5. Обновляем итоги
    Call ОбновитьИтоги(wsData)
    
    ' 6. Автоподбор ширины
    wsGraph.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "График отпусков обновлен!", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

Private Sub ОчиститьОтпуска(ws As Worksheet)
    Dim lastCol As Long
    Dim lastRow As Long
    Dim i As Long, j As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastCol > 2 Then
        For i = 5 To lastRow ' Строки сотрудников начинаются с 5
            For j = 3 To lastCol
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

Private Sub ОбработатьОтпускСотрудника(wsData As Worksheet, wsGraph As Worksheet, _
                                      dataRow As Long, startCol As Long, _
                                      endCol As Long, daysCol As Long, _
                                      graphRow As Long, color As Long)
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim foundCol As Long
    Dim daysCount As Long
    
    On Error Resume Next
    startDate = CDate(wsData.Cells(dataRow, startCol).Value)
    endDate = CDate(wsData.Cells(dataRow, endCol).Value)
    On Error GoTo 0
    
    If IsDate(startDate) And IsDate(endDate) Then
        If endDate >= startDate Then
            ' Считаем дни
            daysCount = DateDiff("d", startDate, endDate) + 1
            wsData.Cells(dataRow, daysCol).Value = daysCount
            
            ' Отмечаем в графике
            currentDate = startDate
            Do While currentDate <= endDate
                foundCol = НайтиСтолбецПоДате(wsGraph, currentDate)
                
                If foundCol > 0 Then
                    With wsGraph.Cells(graphRow, foundCol)
                        .Value = "О"
                        .Interior.Color = color
                        .Font.Bold = True
                        .Font.Color = RGB(0, 100, 0)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                End If
                
                currentDate = DateAdd("d", 1, currentDate)
            Loop
        Else
            wsData.Cells(dataRow, daysCol).Value = "Ошибка"
        End If
    Else
        wsData.Cells(dataRow, daysCol).ClearContents
    End If
End Sub

Private Function НайтиСтолбецПоДате(ws As Worksheet, searchDate As Date) As Long
    Dim col As Long
    Dim lastCol As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 3 To lastCol
        ' Проверяем скрытую строку 4 с полными датами
        If ws.Cells(4, col).Value <> "" Then
            If IsDate(ws.Cells(4, col).Value) Then
                Dim cellDate As Date
                cellDate = CDate(ws.Cells(4, col).Value)
                
                If Year(cellDate) = Year(searchDate) And _
                   Month(cellDate) = Month(searchDate) And _
                   Day(cellDate) = Day(searchDate) Then
                    НайтиСтолбецПоДате = col
                    Exit Function
                End If
            End If
        End If
    Next col
    
    НайтиСтолбецПоДате = 0
End Function

Private Sub ОбновитьИтоги(ws As Worksheet)
    Dim lastRow As Long
    Dim totalDays As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    totalDays = 0
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, 5).Value) Then totalDays = totalDays + ws.Cells(i, 5).Value
        If IsNumeric(ws.Cells(i, 8).Value) Then totalDays = totalDays + ws.Cells(i, 8).Value
        If IsNumeric(ws.Cells(i, 11).Value) Then totalDays = totalDays + ws.Cells(i, 11).Value
    Next i
    
    ws.Cells(lastRow + 1, 1).Value = "ИТОГО дней отпуска:"
    ws.Cells(lastRow + 1, 1).Font.Bold = True
    
    ws.Cells(lastRow + 1, 5).Value = totalDays
    ws.Cells(lastRow + 1, 5).Font.Bold = True
    ws.Cells(lastRow + 1, 5).HorizontalAlignment = xlRight
End Sub

Public Sub Тест()
    MsgBox "Макрос работает! Запустите 'ОбновитьГрафик'", vbInformation
End Sub
'''
    
    # Сохраняем макрос
    macro_file = "макрос_график_отпусков.txt"
    with open(macro_file, "w", encoding="utf-8") as f:
        f.write(macro_code)
    
    print(f"📄 Создан файл с макросом: {macro_file}")
    
    # 10. ИНФОРМАЦИЯ
    print("\n" + "=" * 70)
    print("✅ ФАЙЛ УСПЕШНО СОЗДАН!")
    print("=" * 70)
    
    print(f"\n📁 СОЗДАННЫЕ ФАЙЛЫ:")
    print(f"   1. {filename} - Excel файл с исправленной структурой")
    print(f"   2. {macro_file} - Макрос VBA для обновления")
    
    print(f"\n🎯 ОСОБЕННОСТИ ЭТОЙ ВЕРСИИ:")
    print(f"   • Календарь начинается с колонки C (без смещений)")
    print(f"   • Числа дней: строка 2")
    print(f"   • Дни недели: строка 3")
    print(f"   • Скрытая строка 4: полные даты для макроса")
    print(f"   • Сотрудники: начиная со строки 5")
    print(f"   • Макрос ищет даты в скрытой строке 4")
    
    print(f"\n🚀 КАК ИСПОЛЬЗОВАТЬ:")
    print(f"   1. Откройте {filename} в Excel")
    print(f"   2. Alt+F11 → Insert → Module")
    print(f"   3. Скопируйте код из {macro_file}")
    print(f"   4. Вставьте в модуль")
    print(f"   5. Alt+F8 → выберите 'ОбновитьГрафик'")
    print(f"   6. Нажмите 'Выполнить'")
    
    return filename

def main():
    try:
        create_vacation_schedule()
        
        print("\n" + "=" * 70)
        print("🎯 СТРУКТУРА ФАЙЛА (исправленная):")
        print("=" * 70)
        print("\nЛист ГРАФИК:")
        print("  Строка 1: Названия месяцев (объединенные)")
        print("  Строка 2: Числа дней (1, 2, 3, ...)")
        print("  Строка 3: Дни недели (Пн, Вт, Ср, ...)")
        print("  Строка 4: Скрытые полные даты (для макроса)")
        print("  Строка 5+: Сотрудники (Иванов и т.д.)")
        
        input("\nНажмите Enter для завершения...")
        
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()