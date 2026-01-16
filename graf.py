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
- Данные сотрудников начинаются с строки 5 (строка 4 - только заголовки столбцов)
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
        """Создание листа СОТРУДНИКИ"""
        print("  Создание листа 'СОТРУДНИКИ'...")
        
        # Заголовки
        headers = ["Табельный номер", "Фамилия И.О."]
        for i in range(1, self.vacation_pairs + 1):
            headers.append(f"Отпуск {i} начало")
            headers.append(f"Отпуск {i} конец")
        
        # Записываем заголовки
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, size=11)
            cell.fill = PatternFill(start_color="E0E0E0", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Заполняем табельные номера
        for row in range(2, self.max_employees + 2):
            ws.cell(row=row, column=1, value=row-1)
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
        
        # Настраиваем ширину столбцов
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 25
        for col in range(3, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 14
        
        # Формат дат
        date_format = 'DD.MM.YYYY'
        for row in range(2, self.max_employees + 2):
            for col in range(3, len(headers) + 1):
                ws.cell(row=row, column=col).number_format = date_format
        
        # Пример данных
        example_data = [
            ["Иванов И.И.", datetime.date(self.year, 6, 1), datetime.date(self.year, 6, 14)],
            ["Петров П.П.", datetime.date(self.year, 7, 15), datetime.date(self.year, 7, 28)],
            ["Сидоров С.С.", datetime.date(self.year, 8, 1), datetime.date(self.year, 8, 14)],
        ]
        
        for i, data in enumerate(example_data):
            row = i + 2
            ws.cell(row=row, column=2, value=data[0])
            if len(data) > 1:
                ws.cell(row=row, column=3, value=data[1])
                ws.cell(row=row, column=4, value=data[2])
        
        # Закрепляем заголовки
        ws.freeze_panes = 'A2'
        
        print("  ✓ Лист 'СОТРУДНИКИ' создан")
    
    def _create_schedule_sheet(self, ws):
        """Создание листа ГРАФИК с чередованием цветов месяцев ТОЛЬКО для области данных"""
        print("  Создание листа 'ГРАФИК'...")
        
        # Настраиваем ширину
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 25
        
        # Заголовки столбцов данных (№ и ФИО) - ТОЛЬКО В СТРОКЕ 4
        ws.cell(row=4, column=1, value="№").font = Font(bold=True)
        ws.cell(row=4, column=2, value="ФИО СОТРУДНИКА").font = Font(bold=True)
        
        # Центрируем заголовки данных
        ws.cell(row=4, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=4, column=2).alignment = Alignment(horizontal="center", vertical="center")
        
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
        month_start_cols = []  # Для хранения начала каждого месяца
        month_color_indices = []  # Для хранения цветовых индексов месяцев
        
        for month_idx, month in enumerate(range(1, 13)):
            days_in_month = calendar.monthrange(self.year, month)[1]
            
            # Сохраняем информацию о начале месяца
            month_start_cols.append(current_col)
            month_color_indices.append(month_idx % 2)
            
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
                        fill_color = colors['выходной']
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
        # ИСПРАВЛЕНИЕ: Сохраняем информацию о цветах месяцев для VBA макроса
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
        for i in range(1, self.max_employees + 1):
            row = i + 4  # Строка 5 и ниже (строка 4 - заголовки данных)
            
            # Номер сотрудника
            num_cell = ws.cell(row=row, column=1, value=i)
            num_cell.alignment = Alignment(horizontal="center", vertical="center")
            num_cell.font = Font(size=10)
            
            # ФИО (заполнится макросом)
            name_cell = ws.cell(row=row, column=2, value=f"Сотрудник {i}")
            name_cell.font = Font(size=10)
            
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
        
        print(f"  ✓ Лист 'ГРАФИК' создан с чередованием цветов месяцев ТОЛЬКО в области данных ({current_col-3} дней)")
    
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
            ("2. Заполните столбец 'Фамилия И.О.'", 11, False, False),
            ("3. Заполните даты отпусков в столбцах 'Отпуск N начало/конец'", 11, False, False),
            ("4. Формат дат: ДД.ММ.ГГГГ (например: 15.06.2026)", 11, False, False),
            ("", 1, False, False),
            ("ШАГ 3: ЗАПУСК ГРАФИКА", 14, True, False),
            ("1. Нажмите Alt+F8 для открытия диалога макросов", 11, False, False),
            ("2. Выберите макрос 'ОбновитьГрафик' → 'Выполнить'", 11, False, False),
            ("3. Перейдите на лист 'ГРАФИК' для просмотра", 11, False, False),
            ("", 1, False, False),
            ("ВАЖНО!", 14, True, False),
            ("• Максимальное количество сотрудников: 20", 11, False, False),
            ("• Максимальное количество периодов отпуска: 10 на сотрудника", 11, False, False),
            ("• При изменении дат отпуска перезапустите макрос", 11, False, False),
            ("• Пустые строки (без ФИО) игнорируются", 11, False, False),
        ]
        
        for i, (text, size, bold, center) in enumerate(content, 1):
            cell = ws.cell(row=i, column=1, value=text)
            cell.font = Font(size=size, bold=bold)
            if center:
                cell.alignment = Alignment(horizontal="center")
        
        print("  ✓ Лист 'ИНСТРУКЦИЯ' создан")
    
    def create_vba_macro_file(self):
        """Создание файла с VBA макросом (ИСПРАВЛЕННЫЙ - сохраняет чередование цветов месяцев)"""
        print("\nСоздание файла с VBA макросом...")
        
        vba_code = '''Option Explicit

Public Const MAX_EMPLOYEES As Integer = 20

' Цвета для чередования месяцев
Private Const COLOR_MONTH_1 As Long = &HF8F8F8     ' Светло-серый
Private Const COLOR_MONTH_2 As Long = &HFFFFFF     ' Белый

' Дни в месяцах 2026 года
Private Const DAYS_IN_MONTHS As String = "31,29,31,30,31,30,31,31,30,31,30,31"

' Цвет для отпуска
Private Const VACATION_COLOR As Long = &HC6EFCE    ' Светло-зеленый

Sub ОбновитьГрафик()
    ' Макрос для обновления графика отпусков
    ' Считывает данные с листа СОТРУДНИКИ и заполняет лист ГРАФИК
    ' СОХРАНЯЕТ чередование цветов месяцев в области данных
    
    Dim wsEmployees As Worksheet
    Dim wsSchedule As Worksheet
    Dim wsService As Worksheet
    
    Dim lastRow As Long
    Dim empRow As Long
    Dim scheduleRow As Long
    Dim dateCol As Long
    
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    
    Dim employeeCount As Long
    Dim vacationCount As Long
    Dim periodCount As Long
    
    Dim i As Long, j As Long
    Dim startCol As Long, endCol As Long
    Dim foundDate As Range
    
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
    
    ' Находим последнюю строку с данными на листе СОТРУДНИКИ
    lastRow = wsEmployees.Cells(wsEmployees.Rows.Count, "B").End(xlUp).Row
    
    ' Ограничиваем обработку максимум 20 сотрудниками
    If lastRow > MAX_EMPLOYEES + 1 Then lastRow = MAX_EMPLOYEES + 1
    
    ' Обрабатываем каждого сотрудника
    For i = 2 To lastRow
        ' Проверяем, есть ли ФИО в строке
        If Trim(wsEmployees.Cells(i, 2).Value) <> "" Then
            employeeCount = employeeCount + 1
            
            ' Данные сотрудников начинаются с СТРОКИ 5 (строка 4 - заголовки)
            scheduleRow = employeeCount + 4
            
            ' Копируем табельный номер и ФИО на лист ГРАФИК
            wsSchedule.Cells(scheduleRow, 1).Value = wsEmployees.Cells(i, 1).Value
            wsSchedule.Cells(scheduleRow, 2).Value = wsEmployees.Cells(i, 2).Value
            
            ' Восстанавливаем чередование цветов месяцев для этой строки
            Call ВосстановитьЦветаМесяцев(wsSchedule, scheduleRow)
            
            ' Обрабатываем периоды отпусков (максимум 10 периодов)
            periodCount = 0
            For j = 1 To 10
                startCol = 2 * j + 1  ' Столбцы: 3, 5, 7, 9, 11, 13, 15, 17, 19, 21
                endCol = startCol + 1
                
                ' Проверяем, есть ли дата начала отпуска
                If Not IsEmpty(wsEmployees.Cells(i, startCol).Value) Then
                    If IsDate(wsEmployees.Cells(i, startCol).Value) Then
                        startDate = CDate(wsEmployees.Cells(i, startCol).Value)
                        
                        ' Проверяем, есть ли дата окончания отпуска
                        If Not IsEmpty(wsEmployees.Cells(i, endCol).Value) Then
                            If IsDate(wsEmployees.Cells(i, endCol).Value) Then
                                endDate = CDate(wsEmployees.Cells(i, endCol).Value)
                                
                                ' Проверяем корректность дат (конец не раньше начала)
                                If endDate >= startDate Then
                                    periodCount = periodCount + 1
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
        
        ' Применяем цвет ко всем дням месяца в указанной строке
        For j = 1 To daysInMonth
            With wsSchedule.Cells(rowNum, col)
                .Interior.Color = monthColor
            End With
            col = col + 1
        Next j
    Next i
End Sub

Sub ТестовыеДанные()
    ' Процедура для заполнения тестовых данных
    
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("СОТРУДНИКИ")
    
    ' Очищаем старые данные (кроме заголовков)
    ws.Range("A2:V21").ClearContents
    
    ' Восстанавливаем номера строк
    For i = 1 To 20
        ws.Cells(i + 1, 1).Value = i
    Next i
    
    ' Заполняем тестовые данные
    ws.Cells(2, 2).Value = "Иванов И.И."
    ws.Cells(2, 3).Value = DateSerial(2026, 6, 1)
    ws.Cells(2, 4).Value = DateSerial(2026, 6, 14)
    
    ws.Cells(3, 2).Value = "Петров П.П."
    ws.Cells(3, 3).Value = DateSerial(2026, 7, 15)
    ws.Cells(3, 4).Value = DateSerial(2026, 7, 28)
    
    ws.Cells(4, 2).Value = "Сидоров С.С."
    ws.Cells(4, 3).Value = DateSerial(2026, 8, 1)
    ws.Cells(4, 4).Value = DateSerial(2026, 8, 14)
    
    ' Форматируем даты
    ws.Range("C2:V21").NumberFormat = "DD.MM.YYYY"
    
    ' Автоподбор ширины столбцов
    ws.Columns("A:V").AutoFit
    
    MsgBox "Тестовые данные успешно добавлены!" & vbCrLf & _
           "Запустите макрос 'ОбновитьГрафик' для построения графика.", _
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
        print("\nВАЖНОЕ ИСПРАВЛЕНИЕ В VBA МАКРОСЕ:")
        print("  • Теперь при обновлении графика сохраняется чередование цветов месяцев")
        print("  • Цвета месяцев восстановлены в процедуре ВосстановитьЦветаМесяцев")
        print("  • Цвет отпуска (светло-зеленый) накладывается ПОВЕРХ цвета месяца")
        print("  • После очистки данных автоматически восстанавливается фон месяцев")
        print("\nЧередование цветов месяцев:")
        print("  • Нечетные месяцы (Янв, Мар, Май, Июл, Сен, Ноя): светло-серый (F8F8F8)")
        print("  • Четные месяцы (Фев, Апр, Июн, Авг, Окт, Дек): белый (FFFFFF)")
        print("=" * 70)
    else:
        print("✗ ОШИБКА! Не удалось создать файлы.")


if __name__ == "__main__":
    main()