"""
Генератор графика отпусков в Excel с VBA макросом для автоматического обновления

ВАЖНЫЙ ПРОМПТ ДЛЯ КОНТЕКСТА ПРЕДЫДУЩИХ ЧАТОВ (должен сохраняться при передаче в новый чат):
--------------------------------------------------------------------------------
Этот Python скрипт создает Excel файл для планирования отпусков сотрудников на 2026 год
с VBA макросом для автоматического обновления графика.

ОСНОВНЫЕ ТРЕБОВАНИЯ:
1. Лист "СОТРУДНИКИ" - таблица для ввода данных:
   - 20 сотрудников (строки 2-21)
   - До 10 периодов отпуска на сотрудника (столбцы C-V: пары начало/конец)
   - Формат дат: ДД.ММ.ГГГГ
2. Лист "ГРАФИК" - визуальное представление календаря на 2026 год:
   - Строки 1-3: заголовки календаря (месяцы, числа, дни недели с символами)
   - Строка 4: заголовки столбцов данных ("№", "ФИО СОТРУДНИКА")
   - Строки 5-24: данные 20 сотрудников
   - Календарь начинается с колонки C
   - Дни отпуска отмечаются буквой "О" на светло-зеленом фоне (RGB(198, 239, 206))
   - Производственный календарь России на 2026 год корректно отображен
   - Чередование цветов месяцев: только заголовки месяцев (строка 1) и область данных (строки 5-24)
3. Лист "ДАТЫ" - служебный скрытый лист с датами для работы макроса
4. Лист "ЛЕГЕНДА" и "ИНСТРУКЦИЯ" - пояснительные листы
5. VBA макрос в отдельном файле .txt для автоматического обновления графика

ИСПРАВЛЕННЫЕ ОШИБКИ:
- Заголовки календаря (строки 1-3) НЕ совпадают с данными сотрудников
- Данные сотрудников начинаются с строкя 5 (строка 4 - только заголовки столбцов)
- VBA макрос корректно рассчитывает строки: scheduleRow = employeeCount + 4
- Очистка графика начинается со строки 5
- Производственный календарь 2026 года исправлен

ВАЖНО: При переносе задачи в новый чат этот промпт должен сохраняться в начале файла
для понимания контекста. Текущие задачи добавляются в конец промпта отдельным блоком.
--------------------------------------------------------------------------------
"""

import os
import datetime
from datetime import timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import calendar

class WorkDay:
    def __init__(self, date, day_type, is_short=False):
        self.date = date
        self.day_type = day_type  # 'рабочий', 'выходной', 'праздник', 'рабочая суббота'
        self.is_short = is_short  # Сокращенный день

class ProductionCalendar:
    def __init__(self, year=2026):
        self.year = year
        self.days = []
        self._generate_calendar()
    
    def _generate_calendar(self):
        """Генерация производственного календаря на 2026 год в России"""
        date = datetime.date(self.year, 1, 1)
        
        # Официальные нерабочие праздничные дни (статья 112 ТК РФ)
        holidays = [
            # Новогодние каникулы и Рождество Христово
            datetime.date(self.year, 1, 1),   # Новый год
            datetime.date(self.year, 1, 2),   # 
            datetime.date(self.year, 1, 3),   # 
            datetime.date(self.year, 1, 4),   # 
            datetime.date(self.year, 1, 5),   # 
            datetime.date(self.year, 1, 6),   # 
            datetime.date(self.year, 1, 7),   # Рождество Христово
            datetime.date(self.year, 1, 8),   # 
            
            # День защитника Отечества
            datetime.date(self.year, 2, 23),  # Понедельник
            
            # Международный женский день
            datetime.date(self.year, 3, 8),   # Воскресенье -> выходной 9 марта
            
            # Праздник Весны и Труда
            datetime.date(self.year, 5, 1),   # Пятница
            
            # День Победы
            datetime.date(self.year, 5, 9),   # Суббота -> выходной 11 мая
            
            # День России
            datetime.date(self.year, 6, 12),  # Пятница
            
            # День народного единства
            datetime.date(self.year, 11, 4),  # Среда
        ]
        
        # Дополнительные выходные дни (переносы с суббот 3 и 10 января)
        extra_holidays = [
            datetime.date(self.year, 1, 9),   # Пятница (перенос с сб 3 января)
            datetime.date(self.year, 3, 9),   # Понедельник (перенос с сб 10 января)
            datetime.date(self.year, 5, 11),  # Понедельник (перенос с сб 9 мая)
        ]
        
        # Все праздничные дни
        all_holidays = holidays + extra_holidays
        
        # Предпраздничные дни (сокращенные на 1 час)
        pre_holidays = [
            datetime.date(self.year, 2, 20),  # Пятница перед 23 февраля
            datetime.date(self.year, 3, 7),   # Суббота перед 8 марта
            datetime.date(self.year, 4, 30),  # Четверг перед 1 мая
            datetime.date(self.year, 5, 8),   # Пятница перед 9 мая
            datetime.date(self.year, 6, 11),  # Пятница перед 12 июня
            datetime.date(self.year, 11, 3),  # Вторник перед 4 ноября
            datetime.date(self.year, 12, 31), # Четверг перед Новым годом
        ]
        
        # Рабочие субботы (перенесенные рабочие дни)
        working_saturdays = [
            datetime.date(self.year, 2, 27),  # Суббота (перенос на пн 9 марта)
            datetime.date(self.year, 5, 2),   # Суббота (перенос на пн 4 мая)
        ]
        
        while date.year == self.year:
            # Определяем тип дня
            day_type = 'рабочий'
            is_short = False
            
            # Проверяем праздники
            if date in all_holidays:
                day_type = 'праздник'
            # Проверяем рабочие субботы
            elif date in working_saturdays:
                day_type = 'рабочая суббота'
            # Проверяем обычные выходные (суббота, воскресенье)
            elif date.weekday() >= 5:  # 5=суббота, 6=воскресенье
                day_type = 'выходной'
            
            # Проверяем предпраздничные дни
            if date in pre_holidays and day_type == 'рабочий':
                is_short = True
            
            self.days.append(WorkDay(date, day_type, is_short))
            date += timedelta(days=1)
    
    def get_day_info(self, date):
        """Получить информацию о дне"""
        for day in self.days:
            if day.date == date:
                return day
        return None

class VacationScheduleGenerator:
    def __init__(self, company_name="ООО РОГА И КОПЫТА"):
        self.company_name = company_name
        self.year = 2026
        self.max_employees = 20
        self.vacation_pairs = 10
        self.calendar = ProductionCalendar(self.year)
    
    def create_excel_file(self):
        """Создание Excel файла"""
        print("Создание файла Excel...")
        
        wb = Workbook()
        
        # Удаляем дефолтный лист
        if 'Sheet' in wb.sheetnames:
            del wb['Sheet']
        
        # Создаем листы в правильном порядке
        ws_employees = wb.create_sheet("СОТРУДНИКИ", 0)
        ws_schedule = wb.create_sheet("ГРАФИК", 1)
        ws_dates = wb.create_sheet("ДАТЫ", 2)  # Служебный лист
        ws_legend = wb.create_sheet("ЛЕГЕНДА", 3)
        ws_instruction = wb.create_sheet("ИНСТРУКЦИЯ", 4)
        
        # Скрываем служебный лист
        ws_dates.sheet_state = 'hidden'
        
        # Заполняем листы
        self._create_employees_sheet(ws_employees)
        self._create_schedule_sheet(ws_schedule)
        self._create_dates_sheet(ws_dates)
        self._create_legend_sheet(ws_legend)
        self._create_instruction_sheet(ws_instruction)
        
        # Сохраняем файл
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"отпуск_{self.company_name}_{self.year}_{current_date}.xlsx"
        
        try:
            wb.save(filename)
            print(f"✓ Файл создан: {filename}")
            return filename
        except Exception as e:
            print(f"✗ Ошибка при сохранении файла: {e}")
            return None
    
    def _create_employees_sheet(self, ws):
        """Создание листа СОТРУДНИКИ с новым форматом (блоки сотрудников по горизонтали)"""
        print("  Создание листа 'СОТРУДНИКИ'...")
        
        # Скрываем сетку для лучшего восприятия
        ws.sheet_view.showGridLines = False
        
        # Ширина столбцов для блоков
        block_width = 4  # колонки на одного сотрудника
        date_width = 12
        days_width = 8
        name_width = 25
        
        # Настройки шрифтов
        header_font = Font(bold=True, size=11, color="FFFFFF")
        data_font = Font(size=10)
        total_font = Font(bold=True, size=10)
        
        # Цвета
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # Синий
        name_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")     # Светло-синий
        days_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")     # Светло-зеленый
        
        # Формат дат
        date_format = 'DD.MM.YYYY'
        
        # Размеры блока
        BLOCK_COLS = 4  # ФИО, Дни всего, Дата начала, Дата конца
        MAX_PERIODS = 10  # Максимальное количество периодов отпуска
        
        for emp_index in range(self.max_employees):
            # Определяем начальную колонку для блока сотрудника
            start_col = emp_index * BLOCK_COLS + 1
            
            # --- ЗАГОЛОВКИ БЛОКА ---
            # Строка 1: Заголовки столбцов
            headers = ["ФИО", "дни всего", "", ""]
            for col_offset, header in enumerate(headers):
                col = start_col + col_offset
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(
                    left=Side(style='thin', color="000000"),
                    right=Side(style='thin', color="000000"),
                    top=Side(style='thin', color="000000"),
                    bottom=Side(style='thin', color="000000")
                )
            
            # Строка 2: Пустая строка (отступ)
            for col_offset in range(BLOCK_COLS):
                col = start_col + col_offset
                cell = ws.cell(row=2, column=col, value="")
                cell.border = Border(
                    left=Side(style='thin', color="000000"),
                    right=Side(style='thin', color="000000"),
                    top=Side(style='thin', color="000000"),
                    bottom=Side(style='thin', color="000000")
                )
            
            # --- ДАННЫЕ СОТРУДНИКА ---
            # Строка 3: ФИО и общее количество дней
            # Ячейка ФИО
            name_cell = ws.cell(row=3, column=start_col, value=f"Сотрудник {emp_index+1}")
            name_cell.font = Font(bold=True, size=10)
            name_cell.fill = name_fill
            name_cell.alignment = Alignment(horizontal="left", vertical="center")
            name_cell.border = Border(
                left=Side(style='thin', color="000000"),
                right=Side(style='thin', color="000000"),
                top=Side(style='thin', color="000000"),
                bottom=Side(style='thin', color="000000")
            )
            
            # Ячейка "дни всего" (пока пустая)
            days_cell = ws.cell(row=3, column=start_col+1, value="")
            days_cell.font = total_font
            days_cell.fill = days_fill
            days_cell.alignment = Alignment(horizontal="center", vertical="center")
            days_cell.border = Border(
                left=Side(style='thin', color="000000"),
                right=Side(style='thin', color="000000"),
                top=Side(style='thin', color="000000"),
                bottom=Side(style='thin', color="000000")
            )
            
            # Пустые ячейки справа
            for col_offset in range(2, BLOCK_COLS):
                col = start_col + col_offset
                cell = ws.cell(row=3, column=col, value="")
                cell.border = Border(
                    left=Side(style='thin', color="000000"),
                    right=Side(style='thin', color="000000"),
                    top=Side(style='thin', color="000000"),
                    bottom=Side(style='thin', color="000000")
                )
            
            # --- ПЕРИОДЫ ОТПУСКОВ ---
            # Периоды начинаются с строки 4
            for period_idx in range(MAX_PERIODS):
                row = 4 + period_idx
                
                # Дата начала отпуска
                start_cell = ws.cell(row=row, column=start_col+2, value="")
                start_cell.number_format = date_format
                start_cell.font = data_font
                start_cell.alignment = Alignment(horizontal="center", vertical="center")
                start_cell.border = Border(
                    left=Side(style='thin', color="000000"),
                    right=Side(style='thin', color="000000"),
                    top=Side(style='thin', color="000000"),
                    bottom=Side(style='thin', color="000000")
                )
                
                # Дата окончания отпуска
                end_cell = ws.cell(row=row, column=start_col+3, value="")
                end_cell.number_format = date_format
                end_cell.font = data_font
                end_cell.alignment = Alignment(horizontal="center", vertical="center")
                end_cell.border = Border(
                    left=Side(style='thin', color="000000"),
                    right=Side(style='thin', color="000000"),
                    top=Side(style='thin', color="000000"),
                    bottom=Side(style='thin', color="000000")
                )
                
                # Количество дней в периоде (будет рассчитываться формулой)
                days_period_cell = ws.cell(row=row, column=start_col+1, value="")
                days_period_cell.font = data_font
                days_period_cell.alignment = Alignment(horizontal="center", vertical="center")
                days_period_cell.border = Border(
                    left=Side(style='thin', color="000000"),
                    right=Side(style='thin', color="000000"),
                    top=Side(style='thin', color="000000"),
                    bottom=Side(style='thin', color="000000")
                )
                
                # Ячейка ФИО для выравнивания (пустая)
                empty_cell = ws.cell(row=row, column=start_col, value="")
                empty_cell.border = Border(
                    left=Side(style='thin', color="000000"),
                    right=Side(style='thin', color="000000"),
                    top=Side(style='thin', color="000000"),
                    bottom=Side(style='thin', color="000000")
                )
            
            # --- ПОДВАЛ БЛОКА ---
            # Последняя строка с многоточиями
            last_row = 4 + MAX_PERIODS
            
            for col_offset in range(BLOCK_COLS):
                col = start_col + col_offset
                cell = ws.cell(row=last_row, column=col, value="...")
                cell.font = Font(size=9, italic=True, color="666666")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(
                    left=Side(style='thin', color="000000"),
                    right=Side(style='thin', color="000000"),
                    top=Side(style='thin', color="000000"),
                    bottom=Side(style='thin', color="000000")
                )
            
            # Настраиваем ширину столбцов
            ws.column_dimensions[get_column_letter(start_col)].width = name_width      # ФИО
            ws.column_dimensions[get_column_letter(start_col+1)].width = days_width    # дни всего
            ws.column_dimensions[get_column_letter(start_col+2)].width = date_width    # дата начала
            ws.column_dimensions[get_column_letter(start_col+3)].width = date_width    # дата окончания
            
            # Добавляем вертикальный отступ между блоками
            if emp_index < self.max_employees - 1:
                separator_col = start_col + BLOCK_COLS
                ws.column_dimensions[get_column_letter(separator_col)].width = 2
        
        # Добавляем формулу для расчета дней в каждом периоде
        self._add_formulas_for_days(ws)
        
        # Добавляем пример данных для первых трех сотрудников
        self._add_example_data(ws)
        
        print(f"  ✓ Лист 'СОТРУДНИКИ' создан с {self.max_employees} блоками сотрудников")
    
    def _add_formulas_for_days(self, ws):
        """Добавляем формулы Excel для расчета количества дней отпуска"""
        # Для каждого сотрудника (20 сотрудников)
        for emp_idx in range(self.max_employees):
            start_col = emp_idx * 4 + 1  # Каждый блок занимает 4 колонки
            
            # Для каждого периода (10 периодов)
            for period_idx in range(self.vacation_pairs):
                row = 4 + period_idx
                
                # Формула для расчета дней в периоде: =ЕСЛИ(И(C5<>"";D5<>"");D5-C5+1;"")
                # Где C5 - дата начала, D5 - дата окончания
                start_col_letter = get_column_letter(start_col + 2)  # Колонка даты начала
                end_col_letter = get_column_letter(start_col + 3)    # Колонка даты окончания
                
                # Ячейка для количества дней в периоде (вторая колонка блока)
                days_cell = ws.cell(row=row, column=start_col+1)
                formula = f'=IF(AND({start_col_letter}{row}<>"",{end_col_letter}{row}<>""),{end_col_letter}{row}-{start_col_letter}{row}+1,"")'
                days_cell.value = formula
                days_cell.number_format = '0'  # Целое число
        
        # Формула для итогового количества дней (строка 3, вторая колонка каждого блока)
        for emp_idx in range(self.max_employees):
            start_col = emp_idx * 4 + 1
            total_row = 3
            total_col = start_col + 1
            
            # Собираем ссылки на все ячейки с днями периодов для этого сотрудника
            period_refs = []
            for period_idx in range(self.vacation_pairs):
                row = 4 + period_idx
                period_refs.append(f'{get_column_letter(start_col+1)}{row}')
            
            # Формула для суммы: =СУММ(B5:B14)
            total_formula = f'=SUM({":".join(period_refs)})'
            total_cell = ws.cell(row=total_row, column=total_col)
            total_cell.value = total_formula
    
    def _add_example_data(self, ws):
        """Добавляем пример данных для первых трех сотрудников"""
        example_data = [
            # Сотрудник 1
            {
                'name': 'Иванов И.И.',
                'periods': [
                    (datetime.date(2026, 1, 10), datetime.date(2026, 1, 20)),  # 11 дней
                    (datetime.date(2026, 3, 20), datetime.date(2026, 3, 25)),  # 6 дней
                    (datetime.date(2026, 6, 15), datetime.date(2026, 6, 20)),  # 6 дней
                ]
            },
            # Сотрудник 2
            {
                'name': 'Петров П.П.',
                'periods': [
                    (datetime.date(2026, 7, 15), datetime.date(2026, 7, 28)),  # 14 дней
                ]
            },
            # Сотрудник 3
            {
                'name': 'Сидоров С.С.',
                'periods': [
                    (datetime.date(2026, 8, 1), datetime.date(2026, 8, 14)),  # 14 дней
                ]
            }
        ]
        
        for emp_idx, data in enumerate(example_data):
            if emp_idx >= 3:  # Только первые 3 сотрудника
                break
                
            start_col = emp_idx * 4 + 1
            
            # ФИО
            ws.cell(row=3, column=start_col, value=data['name'])
            
            # Периоды отпусков
            for period_idx, (start_date, end_date) in enumerate(data['periods']):
                if period_idx >= self.vacation_pairs:
                    break
                    
                row = 4 + period_idx
                ws.cell(row=row, column=start_col+2, value=start_date)  # Дата начала
                ws.cell(row=row, column=start_col+3, value=end_date)    # Дата окончания
    
    def _create_schedule_sheet(self, ws):
        """Создание листа ГРАФИК с исправленным отображением сотрудников"""
        print("  Создание листа 'ГРАФИК'...")
        
        # Настраиваем ширину
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 25
        
        # Заголовки столбцов данных - КОРРЕКТИРУЕМ
        ws.cell(row=4, column=1, value="№").font = Font(bold=True)
        ws.cell(row=4, column=2, value="ФИО СОТРУДНИКА").font = Font(bold=True)
        
        # Центрируем заголовки данных
        ws.cell(row=4, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=4, column=2).alignment = Alignment(horizontal="left", vertical="center")
        
        # Заливка для заголовков данных
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        ws.cell(row=4, column=1).fill = header_fill
        ws.cell(row=4, column=2).fill = header_fill
        
        # Создаем календарь на 2026 год - НАЧИНАЕМ С КОЛОНКИ C
        current_col = 3  # Колонка C - первый день года
        
        month_names = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 
                      'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
        
        day_names = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
        
        # Цвета для разных типов дней
        colors = {
            'рабочий': 'FFFFFF',      # Белый
            'выходной': 'F2F2F2',     # Серый
            'праздник': 'FF9999',     # Красный
            'рабочая суббота': '99FF99', # Зеленый
        }
        
        # Цвета для чередования месяцев (светло-серый и белый)
        month_colors = ['F8F8F8', 'FFFFFF']  # Чередующиеся цвета для месяцев
        
        # Создаем заголовки календаря (строки 1-3)
        for month_idx, month in enumerate(range(1, 13)):
            days_in_month = calendar.monthrange(self.year, month)[1]
            
            # Объединяем ячейки для названия месяца (СТРОКА 1)
            start_col = current_col
            end_col = current_col + days_in_month - 1
            
            ws.merge_cells(
                start_row=1, start_column=start_col,
                end_row=1, end_column=end_col
            )
            
            # ТОЛЬКО заголовок месяца получает чередующийся цвет
            month_color = month_colors[month_idx % 2]  # Чередование цветов
            month_fill = PatternFill(start_color=month_color, end_color=month_color, fill_type="solid")
            
            month_cell = ws.cell(row=1, column=start_col, value=month_names[month-1])
            month_cell.alignment = Alignment(horizontal="center", vertical="center")
            month_cell.font = Font(bold=True, size=11)
            month_cell.fill = month_fill  # Только заголовок месяца получает чередующийся цвет
            
            # Заполняем числа (СТРОКА 2) и дни недели (СТРОКА 3) - БЕЗ чередующегося цвета
            for day in range(1, days_in_month + 1):
                col = current_col + day - 1
                date_obj = datetime.date(self.year, month, day)
                day_info = self.calendar.get_day_info(date_obj)
                
                # Число месяца (СТРОКА 2) - БЕЗ чередующегося цвета месяца
                day_cell = ws.cell(row=2, column=col, value=day)
                day_cell.alignment = Alignment(horizontal="center", vertical="center")
                day_cell.font = Font(size=9)
                
                # День недели + символ (СТРОКА 3) - БЕЗ чередующегося цвета месяца
                day_name = day_names[date_obj.weekday()]
                
                # Добавляем символ для особых дней согласно производственному календарю
                symbol = ""
                if day_info:
                    if day_info.day_type == 'праздник':
                        symbol = " ✶"  # Праздник
                    elif day_info.is_short:
                        symbol = " ●"  # Сокращенный (предпраздничный)
                    elif day_info.day_type == 'рабочая суббота':
                        symbol = " ◉"  # Рабочая суббота
                    elif day_info.day_type == 'выходной' and date_obj.weekday() < 5:
                        symbol = " ✶"  # Дополнительный выходной (перенос)
                
                weekday_cell = ws.cell(row=3, column=col, value=f"{day_name}{symbol}")
                weekday_cell.alignment = Alignment(horizontal="center", vertical="center")
                weekday_cell.font = Font(size=8)
                
                # Устанавливаем цвет фона для строк 2-3 согласно типу дня
                fill_color = colors['рабочий']  # По умолчанию белый
                
                if day_info:
                    if day_info.day_type == 'праздник':
                        fill_color = colors['праздник']
                    elif day_info.day_type == 'выходной':
                        fill_color = colors['выходinal']
                    elif day_info.day_type == 'рабочая суббота':
                        fill_color = colors['рабочая суббота']
                    # Для рабочих дней оставляем белый цвет
                
                day_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                day_cell.fill = day_fill
                weekday_cell.fill = day_fill
                
                # Узкие колонки для дней
                ws.column_dimensions[get_column_letter(col)].width = 3.5
            
            current_col += days_in_month
        
        # Применяем чередование цветов ТОЛЬКО к области данных (строки 5+) 
        data_start_row = 5  # Данные сотрудников начинаются с строки 5
        data_end_row = data_start_row + self.max_employees - 1  # Строка 24
        
        # Применяем цвет месяца ко всем ячейкам этого месяца в строках данных
        col = 3  # Начинаем с колонки C
        for month_idx, month in enumerate(range(1, 13)):
            days_in_month = calendar.monthrange(self.year, month)[1]
            
            # Получаем цвет для этого месяца
            month_color = month_colors[month_idx % 2]
            month_fill = PatternFill(start_color=month_color, end_color=month_color, fill_type="solid")
            
            # Применяем цвет месяца ко всем ячейкам этого месяца в строках данных
            for day in range(1, days_in_month + 1):
                for row in range(data_start_row, data_end_row + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = month_fill
                col += 1
        
        # Добавляем строки для сотрудников (начинаем с СТРОКИ 5)
        # ТОЛЬКО номера сотрудников - ФИО будут заполняться макросом из листа СОТРУДНИКИ
        for i in range(1, self.max_employees + 1):
            row = i + 4  # Строка 5 и ниже (строка 4 - заголовки данных)
            
            # Номер сотрудника
            num_cell = ws.cell(row=row, column=1, value=i)
            num_cell.alignment = Alignment(horizontal="center", vertical="center")
            num_cell.font = Font(size=10)
            
            # Ячейка для ФИО (оставляем пустой - заполнится макросом)
            name_cell = ws.cell(row=row, column=2)
            name_cell.font = Font(size=10)
            name_cell.alignment = Alignment(horizontal="left", vertical="center")
            
            # Форматируем строки данных
            for col in range(1, current_col):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                cell.alignment = Alignment(vertical="center", horizontal="center")
            # Для столбца с ФИО выравнивание слева
            ws.cell(row=row, column=2).alignment = Alignment(vertical="center", horizontal="left")
        
        # Закрепляем области: строка 4 (заголовки данных) и колонки A,B
        ws.freeze_panes = 'C5'
        
        # Добавляем итоговую информацию
        last_col_letter = get_column_letter(current_col - 1)
        footer_row = self.max_employees + 6
        ws.merge_cells(f'A{footer_row}:{last_col_letter}{footer_row}')
        footer = ws.cell(row=footer_row, column=1, 
                        value=f"График отпусков {self.company_name} на {self.year} год")
        footer.font = Font(italic=True, size=10, color="666666")
        footer.alignment = Alignment(horizontal="center")
        
        print(f"  ✓ Лист 'ГРАФИК' создан ({current_col-3} дней)")
    
    def _create_dates_sheet(self, ws):
        """Создание служебного листа с датами"""
        print("  Создание служебного листа с датами...")
        
        # Заполняем даты по горизонтали
        date_obj = datetime.date(self.year, 1, 1)
        col = 1
        
        while date_obj.year == self.year:
            cell = ws.cell(row=1, column=col, value=date_obj)
            cell.number_format = 'DD.MM.YYYY'
            date_obj += timedelta(days=1)
            col += 1
        
        # Минимизируем видимость
        for c in range(1, col):
            ws.column_dimensions[get_column_letter(c)].width = 0.5
        
        print(f"  ✓ Служебный лист создан ({col-1} дней)")
    
    def _create_legend_sheet(self, ws):
        """Создание листа ЛЕГЕНДА"""
        print("  Создание листа 'ЛЕГЕНДА'...")
        
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 40
        
        # Заголовок
        ws.merge_cells('A1:C1')
        title = ws.cell(row=1, column=1, value="ЛЕГЕНДА ГРАФИКА ОТПУСКОВ")
        title.font = Font(bold=True, size=14, color="1F4E78")
        title.alignment = Alignment(horizontal="center")
        
        # Заголовки таблицы
        headers = ["Обозначение", "Тип дня", "Описание"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Данные легенды
        legend_data = [
            ("✶", "Праздник/Выходной", "Нерабочий праздничный или дополнительный выходной день", "FF9999"),
            ("●", "Сокращенный", "Предпраздничный рабочий день (короче на 1 час)", "FFFF99"),
            ("◉", "Рабочая суббота", "Перенесенный рабочий день (суббота)", "99FF99"),
            ("О", "Отпуск", "День отпуска сотрудника", "C6EFCE"),
            ("", "Выходной", "Суббота, воскресенье", "F2F2F2"),
            ("", "Рабочий", "Обычный рабочий день", "FFFFFF"),
        ]
        
        for i, (symbol, day_type, description, color) in enumerate(legend_data, 1):
            row = i + 3
            
            # Символ
            sym_cell = ws.cell(row=row, column=1, value=symbol)
            sym_cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Тип дня
            type_cell = ws.cell(row=row, column=2, value=day_type)
            
            # Описание
            desc_cell = ws.cell(row=row, column=3, value=description)
            
            # Цвет фона
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            sym_cell.fill = fill
            type_cell.fill = fill
            desc_cell.fill = fill
        
        print("  ✓ Лист 'ЛЕГЕНДА' создан")
    
    def _create_instruction_sheet(self, ws):
        """Создание листа ИНСТРУКЦИЯ"""
        print("  Создание листа 'ИНСТРУКЦИЯ'...")
        
        ws.column_dimensions['A'].width = 80
        
        content = [
            ("ИНСТРУКЦИЯ ПО РАБОТЕ С ГРАФИКОМ ОТПУСКОВ", 16, True, True),
            ("", 1, False, False),
            ("ШАГ 1: УСТАНОВКА МАКРОСА", 14, True, False),
            ("1. Откройте файл в Microsoft Excel", 11, False, False),
            ("2. Нажмите Alt+F11 для открытия редактора VBA", 11, False, False),
            ("3. В меню выберите Insert → Module", 11, False, False),
            ("4. Скопируйте код из файла vacation_macro.txt в новый модуль", 11, False, False),
            ("5. Закройте редактор VBA (Alt+Q)", 11, False, False),
            ("", 1, False, False),
            ("ШАГ 2: ЗАПОЛНЕНИЕ ДАННЫХ", 14, True, False),
            ("1. Перейдите на лист 'СОТРУДНИКИ'", 11, False, False),
            ("2. Заполните столбец 'ФИО' в каждом блоке сотрудника", 11, False, False),
            ("3. Заполните даты отпусков в столбцах 'Дата начала' и 'Дата окончания'", 11, False, False),
            ("4. Формат дат: ДД.ММ.ГГГГ (например: 15.06.2026)", 11, False, False),
            ("5. Количество дней рассчитывается автоматически", 11, False, False),
            ("", 1, False, False),
            ("ШАГ 3: ЗАПУСК ГРАФИКА", 14, True, False),
            ("1. Нажмите Alt+F8 для открытия диалога макросов", 11, False, False),
            ("2. Выберите макрос 'ОбновитьГрафик' → 'Выполнить'", 11, False, False),
            ("3. Перейдите на лист 'ГРАФИК' для просмотра", 11, False, False),
            ("", 1, False, False),
            ("ВАЖНОЕ ИЗМЕНЕНИЕ В ФОРМАТЕ!", 14, True, False),
            ("• Теперь каждый сотрудник имеет отдельный блок из 4 колонок", 11, False, False),
            ("• Структура блока: ФИО | дни всего | Дата начала | Дата окончания", 11, False, False),
            ("• До 10 периодов отпуска на сотрудника", 11, False, False),
            ("• Количество дней рассчитывается автоматически формулой", 11, False, False),
            ("• Формула: =ЕСЛИ(И(дата_начала<>"";дата_конца<>"");дата_конца-дата_начала+1;"")", 11, False, False),
            ("• Итоговые дни: =СУММ(диапазон_дней_периодов)", 11, False, False),
        ]
        
        for i, (text, size, bold, center) in enumerate(content, 1):
            cell = ws.cell(row=i, column=1, value=text)
            cell.font = Font(size=size, bold=bold)
            if center:
                cell.alignment = Alignment(horizontal="center")
        
        print("  ✓ Лист 'ИНСТРУКЦИЯ' создан")
    
    def create_vba_macro_file(self):
        """Создание файла с VBA макросом для нового формата листа СОТРУДНИКИ"""
        print("\nСоздание файла с VBA макросом...")
        
        vba_code = '''Option Explicit

Public Const MAX_EMPLOYEES As Integer = 20
Public Const BLOCK_COLS As Integer = 4         ' Колонок на одного сотрудника (ФИО, дни, начало, конец)
Public Const MAX_PERIODS As Integer = 10       ' Максимальное количество периодов отпуска

' Цвета для чередования месяцев
Private Const COLOR_MONTH_1 As Long = &HF8F8F8     ' Светло-серый
Private Const COLOR_MONTH_2 As Long = &HFFFFFF     ' Белый

' Дни в месяцах 2026 года
Private Const DAYS_IN_MONTHS As String = "31,29,31,30,31,30,31,31,30,31,30,31"

' Цвет для отпуска
Private Const VACATION_COLOR As Long = &HC6EFCE    ' Светло-зеленый

Sub ОбновитьГрафик()
    ' Макрос для обновления графика отпусков
    ' Считывает данные с листа СОТРУДНИКИ (новый формат) и заполняет лист ГРАФИК
    
    Dim wsEmployees As Worksheet
    Dim wsSchedule As Worksheet
    Dim wsService As Worksheet
    
    Dim employeeCount As Long
    Dim vacationCount As Long
    Dim i As Long, j As Long
    Dim empBlockStart As Long
    
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim dateCol As Long
    Dim foundDate As Range
    
    Dim scheduleRow As Long
    Dim nameCell As Range
    Dim startDateCell As Range
    Dim endDateCell As Range
    
    ' Отключаем обновление экрана для скорости
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Инициализация счетчиков
    employeeCount = 0
    vacationCount = 0
    
    ' Получаем ссылки на рабочие листы
    Set wsEmployees = ThisWorkbook.Worksheets("СОТРУДНИКИ")
    Set wsSchedule = ThisWorkbook.Worksheets("ГРАФИК")
    Set wsService = ThisWorkbook.Worksheets("ДАТЫ")
    
    ' Очищаем предыдущий график (НО сохраняем цвет фона)
    Call ОчиститьГрафикСФорматированием(wsSchedule)
    
    ' Обрабатываем каждого сотрудника (максимум 20)
    For i = 0 To MAX_EMPLOYEES - 1
        ' Определяем начало блока сотрудника
        empBlockStart = i * BLOCK_COLS + 1  ' A=1, E=5, I=9, M=13, Q=17, U=21 и т.д.
        
        ' Ячейка с ФИО (строка 3, первая колонка блока)
        Set nameCell = wsEmployees.Cells(3, empBlockStart)
        
        ' Проверяем, есть ли ФИО в ячейке
        If Trim(nameCell.Value) <> "" Then
            employeeCount = employeeCount + 1
            
            ' Данные сотрудников начинаются с СТРОКИ 5 (строка 4 - заголовки)
            scheduleRow = employeeCount + 4
            
            ' Заполняем номер и ФИО на листе ГРАФИК
            wsSchedule.Cells(scheduleRow, 1).Value = employeeCount
            wsSchedule.Cells(scheduleRow, 2).Value = nameCell.Value
            
            ' Восстанавливаем чередование цветов месяцев для этой строки
            Call ВосстановитьЦветаМесяцев(wsSchedule, scheduleRow)
            
            ' Обрабатываем периоды отпусков для этого сотрудника
            For j = 0 To MAX_PERIODS - 1
                ' Строка для периода (начинается с строки 4)
                Dim periodRow As Long
                periodRow = 4 + j
                
                ' Ячейки с датами начала и окончания отпуска
                Set startDateCell = wsEmployees.Cells(periodRow, empBlockStart + 2)  ' Третья колонка блока
                Set endDateCell = wsEmployees.Cells(periodRow, empBlockStart + 3)    ' Четвертая колонка блока
                
                ' Проверяем, есть ли обе даты
                If Not IsEmpty(startDateCell.Value) And Not IsEmpty(endDateCell.Value) Then
                    If IsDate(startDateCell.Value) And IsDate(endDateCell.Value) Then
                        startDate = CDate(startDateCell.Value)
                        endDate = CDate(endDateCell.Value)
                        
                        ' Проверяем корректность дат (конец не раньше начала)
                        If endDate >= startDate Then
                            vacationCount = vacationCount + 1
                            
                            ' Отмечаем все дни отпуска на графике
                            currentDate = startDate
                            Do While currentDate <= endDate
                                ' Ищем столбец с этой датой на служебном листе
                                Set foundDate = wsService.Rows(1).Find( _
                                    What:=currentDate, _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlNext)
                                
                                If Not foundDate Is Nothing Then
                                    dateCol = foundDate.Column
                                    
                                    ' Заполняем ячейку на графике (ЦВЕТ НАКЛАДЫВАЕТСЯ ПОВЕРХ цвета месяца)
                                    With wsSchedule.Cells(scheduleRow, dateCol)
                                        .Value = "О"  ' Буква О - отпуск
                                        .Interior.Color = VACATION_COLOR  ' Светло-зеленый
                                        .Font.Bold = True
                                        .Font.Name = "Arial"
                                        .Font.Size = 9
                                        .HorizontalAlignment = xlCenter
                                        .VerticalAlignment = xlCenter
                                    End With
                                End If
                                
                                currentDate = currentDate + 1
                            Loop
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
    ' Автоподбор ширины столбца с ФИО
    wsSchedule.Columns("B:B").AutoFit
    
    ' Восстанавливаем настройки Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Показываем информационное сообщение
    MsgBox "График отпусков успешно обновлен!" & vbCrLf & _
           "Обработано сотрудников: " & employeeCount & vbCrLf & _
           "Найдено периодов отпуска: " & vacationCount & vbCrLf & _
           "Чередование цветов месяцев сохранено.", _
           vbInformation + vbOKOnly, _
           "Обновление графика отпусков"
    
    ' Активируем лист с графиком
    wsSchedule.Activate
    wsSchedule.Range("A1").Select
    
    Exit Sub

ErrorHandler:
    ' Восстанавливаем настройки при ошибке
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Показываем сообщение об ошибке
    MsgBox "Произошла ошибка при обновления графика!" & vbCrLf & _
           "Код ошибки: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description & vbCrLf & _
           "Проверьте:" & vbCrLf & _
           "1. Корректность формата дат" & vbCrLf & _
           "2. Наличие ФИО в строках" & vbCrLf & _
           "3. Что дата окончания не раньше даты начала", _
           vbCritical + vbOKOnly, _
           "Ошибка обновления графика"
End Sub

Private Sub ОчиститьГрафикСФорматированием(wsSchedule As Worksheet)
    ' Очистка данных на листе ГРАФИК (сохраняет цвета месяцев)
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    
    ' Находим последнюю строку с данными (начинаем поиск со строки 5)
    lastRow = wsSchedule.Cells(wsSchedule.Rows.Count, "B").End(xlUp).Row
    
    ' Если есть данные для очистки (данные начинаются с строки 5)
    If lastRow >= 5 Then
        ' Очищаем номера и ФИО сотрудников (строки 5 и ниже)
        wsSchedule.Range("A5:B" & lastRow).ClearContents
        
        ' Находим последний столбец с датами
        lastCol = wsSchedule.Cells(3, wsSchedule.Columns.Count).End(xlToLeft).Column
        
        ' Очищаем отметки об отпусках (начиная с колонки C, строки 5 и ниже)
        If lastCol >= 3 Then
            ' Очищаем содержимое, НО НЕ форматирование
            For i = 5 To lastRow
                For j = 3 To lastCol
                    With wsSchedule.Cells(i, j)
                        .Value = ""  ' Очищаем только значение
                        .Font.Bold = False
                        .Interior.Pattern = xlNone  ' Убираем цвет отпуска
                    End With
                Next j
            Next i
            
            ' Восстанавливаем цвета месяцев для всех строк данных
            For i = 5 To lastRow
                Call ВосстановитьЦветаМесяцев(wsSchedule, i)
            Next i
        End If
    End If
End Sub

Private Sub ВосстановитьЦветаМесяцев(wsSchedule As Worksheet, rowNum As Long)
    ' Восстанавливает чередование цветов месяцев для указанной строки
    
    Dim monthDays() As String
    Dim i As Long, j As Long
    Dim col As Long
    Dim monthColor As Long
    Dim daysInMonth As Integer
    
    ' Разбиваем строку с днями в месяцах
    monthDays = Split(DAYS_IN_MONTHS, ",")
    
    ' Начинаем с колонки C (3)
    col = 3
    
    ' Проходим по всем месяцам
    For i = 0 To UBound(monthDays)
        daysInMonth = CInt(monthDays(i))
        
        ' Определяем цвет для месяца (чередование)
        If (i Mod 2) = 0 Then
            monthColor = COLOR_MONTH_1  ' Нечетные месяцы: светло-серый
        Else
            monthColor = COLOR_MONTH_2  ' Четные месяцы: белый
        End If
        
        ' Применяем цвет ко всем дням месяца в указанной строки
        For j = 1 To daysInMonth
            With wsSchedule.Cells(rowNum, col)
                .Interior.Color = monthColor
            End With
            col = col + 1
        Next j
    Next i
End Sub

Sub ТестовыеДанные()
    ' Процедура для заполнения тестовых данных (исправлена для нового формата)
    
    Dim ws As Worksheet
    Dim i As Long
    Dim blockStart As Long
    
    Set ws = ThisWorkbook.Worksheets("СОТРУДНИКИ")
    
    ' Очищаем старые данные (сохраняем заголовки и форматирование)
    For i = 0 To MAX_EMPLOYEES - 1
        blockStart = i * BLOCK_COLS + 1
        
        ' Очищаем данные сотрудника (сохраняем формулы)
        ws.Cells(3, blockStart).ClearContents              ' ФИО
        ws.Range(ws.Cells(4, blockStart + 2), ws.Cells(13, blockStart + 3)).ClearContents  ' Даты отпусков
    Next i
    
    ' Заполняем тестовые данные для первых 3 сотрудников
    
    ' Сотрудник 1 (блок A-D)
    ws.Cells(3, 1).Value = "Иванов И.И."
    ws.Cells(4, 3).Value = DateSerial(2026, 1, 10)   ' Начало отпуска 1
    ws.Cells(4, 4).Value = DateSerial(2026, 1, 20)   ' Конец отпуска 1 (11 дней)
    ws.Cells(5, 3).Value = DateSerial(2026, 3, 20)   ' Начало отпуска 2
    ws.Cells(5, 4).Value = DateSerial(2026, 3, 25)   ' Конец отпуска 2 (6 дней)
    ws.Cells(6, 3).Value = DateSerial(2026, 6, 15)   ' Начало отпуска 3
    ws.Cells(6, 4).Value = DateSerial(2026, 6, 20)   ' Конец отпуска 3 (6 дней)
    
    ' Сотрудник 2 (блок E-H)
    ws.Cells(3, 5).Value = "Петров П.П."
    ws.Cells(4, 7).Value = DateSerial(2026, 7, 15)   ' Начало отпуска
    ws.Cells(4, 8).Value = DateSerial(2026, 7, 28)   ' Конец отпуска (14 дней)
    
    ' Сотрудник 3 (блок I-L)
    ws.Cells(3, 9).Value = "Сидоров С.С."
    ws.Cells(4, 11).Value = DateSerial(2026, 8, 1)   ' Начало отпуска
    ws.Cells(4, 12).Value = DateSerial(2026, 8, 14)   ' Конец отпуска (14 дней)
    
    ' Форматируем даты
    For i = 0 To MAX_EMPLOYEES - 1
        blockStart = i * BLOCK_COLS + 1
        ws.Range(ws.Cells(4, blockStart + 2), ws.Cells(13, blockStart + 3)).NumberFormat = "DD.MM.YYYY"
    Next i
    
    MsgBox "Тестовые данные успешно добавлены!" & vbCrLf & _
           "Запустите макрос 'ОбновитьГрафик' для построения графика." & vbCrLf & _
           "Примечание: данные добавлены для первых 3 сотрудников.", _
           vbInformation, "Тестовые данные"
End Sub
'''
        
        # Сохраняем макрос в файл
        filename = "vacation_macro.txt"
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(vba_code)
            print(f"✓ Файл с макросом создан: {filename}")
            return filename
        except Exception as e:
            print(f"✗ Ошибка при создании файла макроса: {e}")
            return None


def main():
    """Основная функция"""
    print("=" * 70)
    print("ГЕНЕРАТОР ГРАФИКА ОТПУСКОВ С VBA МАКРОСОМ")
    print(f"ПРОИЗВОДСТВЕННЫЙ КАЛЕНДАРЬ РОССИИ НА 2026 ГОД")
    print("=" * 70)
    
    # Получаем название компании
    company_name = input("\nВведите название компании: ").strip()
    if not company_name:
        company_name = "ООО РОГА И КОПЫТА"
    
    print("\n" + "=" * 70)
    print("НАЧАЛО СОЗДАНИЯ ФАЙЛОВ...")
    print("=" * 70)
    
    # Создаем генератор
    generator = VacationScheduleGenerator(company_name)
    
    # Создаем файлы
    excel_file = generator.create_excel_file()
    macro_file = generator.create_vba_macro_file()
    
    print("\n" + "=" * 70)
    
    if excel_file and macro_file:
        print("✓ ФАЙЛЫ УСПЕШНО СОЗДАНЫ")
        print(f"  • Excel файл: {excel_file}")
        print(f"  • Файл макроса: {macro_file}")
        print("\nВАЖНЫЕ ИЗМЕНЕНИЯ:")
        print("  1. ЛИСТ 'СОТРУДНИКИ' - новый формат:")
        print("     • Каждый сотрудник имеет отдельный блок из 4 колонок")
        print("     • Блоки расположены по горизонтали")
        print("     • Структура блока: ФИО | дни всего | Дата начала | Дата окончания")
        print("     • До 10 периодов отпуска на сотрудника")
        print("     • Количество дней рассчитывается автоматически формулой")
        print("")
        print("  2. ЛИСТ 'ГРАФИК' - исправленный формат:")
        print("     • В столбце '№' отображаются номера 1-20")
        print("     • В столбце 'ФИО СОТРУДНИКА' будут отображаться ФИО из листа СОТРУДНИКИ")
        print("     • НЕ отображается количество дней отпуска")
        print("")
        print("  3. VBA МАКРОС - полностью переработан:")
        print("     • Читает данные из нового формата листа СОТРУДНИКИ")
        print("     • Правильно определяет блоки сотрудников (по 4 колонки)")
        print("     • Сохраняет чередование цветов месяцев")
        print("     • Цвет отпуска накладывается поверх цвета месяца")
        print("=" * 70)
    else:
        print("✗ ОШИБКА! Не удалось создать файлы.")


if __name__ == "__main__":
    main()