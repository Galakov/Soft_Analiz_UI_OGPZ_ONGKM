import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import sys
# from openpyxl import load_workbook
# from openpyxl.formatting.rule import ColorScaleRule
# from openpyxl.styles import PatternFill, Border, Side, Alignment
import numpy as np
import logging
import unicodedata

def resource_path(relative_path):
    """Получает абсолютный путь к ресурсу, работает для dev, PyInstaller и pip install"""
    try:
        # PyInstaller создает временную папку и хранит путь в _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Normal python script or installed package
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    path = os.path.join(base_path, relative_path)
    # print(f"Debug: Looking for resource at {path}") # Uncomment for debugging
    return path

def get_column_letter(col_idx):
    """Преобразует числовой индекс столбца в буквенное обозначение Excel"""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def format_data_workbook(writer, sheet_name, df, rules_file):
    """
    Форматирует лист с данными используя xlsxwriter (в один проход).
    Добавляет заголовки, объединяет ячейки параметров, настраивает ширину и цвета.
    """
    try:
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Стили
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1,
            'bg_color': '#D9D9D9'
        })
        
        center_format = workbook.add_format({
            'align': 'center', 
            'valign': 'vcenter'
        })
        
        # 1. Подготовка данных для заголовков
        # Читаем правила
        try:
            rules_df = pd.read_excel(rules_file, engine='openpyxl')
        except:
            rules_df = pd.DataFrame()

        # Создаем маппинг
        column_to_node = {}
        param_ranges = {}
        
        for _, row in rules_df.iterrows():
            if len(row) >= 4:
                new_name = str(row.iloc[2]).strip()
                node_name = str(row.iloc[3]).strip()
                if new_name and node_name:
                    column_to_node[new_name] = node_name
                
                # Попытка получить диапазоны
                if len(row) >= 7 and new_name:
                    try:
                        min_val = row.iloc[5]
                        max_val = row.iloc[6]
                        units = str(row.iloc[7]).strip() if len(row) >= 8 and pd.notna(row.iloc[7]) else ""
                        
                        if not units:
                            param_name = str(row.iloc[4]).strip() if len(row) >= 5 else ""
                            check_name = param_name.lower() if param_name else new_name.lower()
                            if "перепад давления" in check_name: units = "кгс/см2"
                            elif "расход" in check_name: units = "тыс. м3/ч"
                            elif "температура" in check_name: units = "°C"
                        
                        if pd.notna(min_val) and pd.notna(max_val):
                            range_str = f"({min_val} ... {max_val} {units})".strip()
                            param_ranges[new_name] = range_str
                    except:
                        pass

        # 2. Запись заголовков (Строки 1 и 2 в Excel -> 0 и 1 индексы)
        headers = df.columns.tolist()
        
        # Группировка для объединения (Строка 0)
        # Ищем последовательные столбцы с одинаковым параметром для range_str
        current_param_range = None
        merge_start_col = 0
        
        # Замораживаем панели
        worksheet.freeze_panes(2, 1) # 2 строки заголовка, 1 столбец слева
        
        for i, header in enumerate(headers):
            # Запись заголовка столбца (Строка 1)
            worksheet.write(1, i, header, header_format)
            
            # Логика объединения для строки 0
            base_header = header.split(' ⚠')[0] if ' ⚠' in header else header
            param_range_str = param_ranges.get(base_header)
            
            if i == 0:
                current_param_range = param_range_str
                merge_start_col = 0
            else:
                if param_range_str != current_param_range:
                    # Записываем предыдущую группу
                    if current_param_range:
                        if i - 1 > merge_start_col:
                            worksheet.merge_range(0, merge_start_col, 0, i - 1, current_param_range, header_format)
                        else:
                            worksheet.write(0, merge_start_col, current_param_range, header_format)
                    # Начинаем новую
                    current_param_range = param_range_str
                    merge_start_col = i
                
        # Записываем последнюю группу
        if current_param_range:
             if len(headers) - 1 > merge_start_col:
                worksheet.merge_range(0, merge_start_col, 0, len(headers) - 1, current_param_range, header_format)
             else:
                worksheet.write(0, merge_start_col, current_param_range, header_format)

        # 3. Настройка ширины столбцов и границ
        # Границы групп узлов
        current_node = None
        node_start_col = 0
        
        thick_border_fmt = workbook.add_format({'left': 2}) # Thick left border
        
        for i, header in enumerate(headers):
            base_header = header.split(' ⚠')[0] if ' ⚠' in header else header
            node = column_to_node.get(base_header)
            
            # Ширина столбца
            max_len = len(str(header))
            # Примерная ширина по данным (первые 50 строк)
            for val in df.iloc[:50, i].astype(str):
                max_len = max(max_len, len(val))
            worksheet.set_column(i, i, min(max_len + 2, 50))
            
            # Условное форматирование (Color Scale) для данных
            if header != 'Время' and not header.endswith('⚠'):
                # Определяем диапазон данных (с 3-й строки Excel, индекс 2)
                # end_row = len(df) + 2 - 1
                # range_str = f"{get_column_letter(i+1)}3:{get_column_letter(i+1)}{len(df)+2}"
                # xlsxwriter conditional format
                worksheet.conditional_format(2, i, len(df)+1, i, {
                    'type': '3_color_scale',
                    'min_color': '#FF0000', # Red
                    'mid_color': '#FFFF00', # Yellow
                    'max_color': '#92D050'  # Green
                })

            # Границы узлов (визуально отделяем группы)
            if node:
                if node != current_node:
                    if i > 0: # Не для первого столбца
                        # Применяем левую границу ко всему столбцу (через cond format hack или просто set_column?)
                        # set_column перетрет ширину.
                        # cond format hack: Always True
                        worksheet.conditional_format(0, i, len(df)+1, i, {
                            'type': 'formula',
                            'criteria': '=TRUE',
                            'format': thick_border_fmt
                        })
                    current_node = node
                    
    except Exception as e:
        print(f"Ошибка при форматировании данных: {e}")
        import traceback
        traceback.print_exc()

def add_arrow_columns(df, rules_file):
    """Добавляет столбцы со стрелками для значений вне диапазона min-max"""
    df = df.copy()
    try:
        # Читаем файл с правилами
        rules_df = pd.read_excel(rules_file, engine='openpyxl')
        
        # Создаем словари для хранения min/max значений параметров и соответствия узлов измерения
        param_limits = {}
        param_to_node = {}  # Словарь для хранения соответствия параметр -> узел измерения
        
        # Заполняем словари из файла правил
        for _, row in rules_df.iterrows():
            if len(row) >= 7:  # Проверяем наличие столбцов min и max
                new_name = str(row.iloc[2]).strip()
                node_name = str(row.iloc[3]).strip()  # Название узла измерения
                min_val = row.iloc[5]  # Предполагаем, что min в 6-м столбце
                max_val = row.iloc[6]  # Предполагаем, что max в 7-м столбце
                
                # Сохраняем соответствие параметра и узла измерения
                if new_name and node_name:
                    param_to_node[new_name] = node_name
                
                # Проверяем, что значения min и max заданы
                if pd.notna(min_val) and pd.notna(max_val) and new_name:
                    try:
                        min_val = float(min_val)
                        max_val = float(max_val)
                        param_limits[new_name] = {'min': min_val, 'max': max_val}
                    except (ValueError, TypeError):
                        print(f"Пропущены некорректные значения min/max для {new_name}")
                        continue
        
        # Добавляем столбцы со стрелками
        for col in df.columns:
            if col in param_limits:
                limits = param_limits[col]
                arrow_col_name = f"{col} ⚠"
                
                # Создаем столбец со стрелками
                df[arrow_col_name] = ''
                
                # Заполняем стрелки на основе условий, пропуская нулевые значения
                numeric_values = pd.to_numeric(df[col], errors='coerce')
                df.loc[(numeric_values < limits['min']) & (numeric_values != 0), arrow_col_name] = '↓'
                df.loc[(numeric_values > limits['max']) & (numeric_values != 0), arrow_col_name] = '↑'
                
                # Перемещаем столбец со стрелками сразу после исходного столбца
                cols = list(df.columns)
                current_idx = cols.index(arrow_col_name)
                target_idx = cols.index(col) + 1
                cols.insert(target_idx, cols.pop(current_idx))
                df = df[cols]
        
        return df, param_to_node
        
    except Exception as e:
        print(f"Ошибка при добавлении столбцов со стрелками: {str(e)}")
        return df, {}

class ExcelMerger:
    def __init__(self, root):
        self.root = root
        self.root.title("Объединение Excel файлов")
        self.root.geometry("1200x600")
        
        # Список для хранения путей к файлам
        self.files = []
        
        # Словари для хранения чекбоксов
        self.parameter_vars = {}
        self.node_vars = {}  # Новый словарь для чекбоксов узлов измерения
        
        # Создание элементов интерфейса
        self.create_widgets()
        
    def create_widgets(self):
        # Создаем фрейм для левой части (файлы)
        left_frame = ttk.Frame(self.root)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Кнопка для добавления файлов
        add_button = ttk.Button(left_frame, text="Добавить файлы", command=self.add_files)
        add_button.pack(pady=10)

        # Кнопка для настройки диапазонов
        ranges_button = ttk.Button(left_frame, text="Настройка диапазонов", command=self.open_range_editor)
        ranges_button.pack(pady=5)
        
        # Список файлов
        self.files_frame = ttk.LabelFrame(left_frame, text="Выбранные файлы")
        self.files_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.files_listbox = tk.Listbox(self.files_frame)
        self.files_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Кнопка удаления выбранного файла
        remove_button = ttk.Button(left_frame, text="Удалить выбранный", command=self.remove_file)
        remove_button.pack(pady=5)
        
        # Фрейм для временного интервала
        time_frame = ttk.LabelFrame(left_frame, text="Временной интервал")
        time_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Начальная дата
        start_frame = ttk.Frame(time_frame)
        start_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(start_frame, text="Начало:").pack(side=tk.LEFT)
        self.start_date = ttk.Entry(start_frame, width=10)
        self.start_date.pack(side=tk.LEFT, padx=5)
        self.start_time = ttk.Entry(start_frame, width=8)
        self.start_time.pack(side=tk.LEFT)
        
        # Конечная дата
        end_frame = ttk.Frame(time_frame)
        end_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(end_frame, text="Конец:").pack(side=tk.LEFT)
        self.end_date = ttk.Entry(end_frame, width=10)
        self.end_date.pack(side=tk.LEFT, padx=5)
        self.end_time = ttk.Entry(end_frame, width=8)
        self.end_time.pack(side=tk.LEFT)
        
        # Подсказка формата
        ttk.Label(time_frame, text="Формат: ГГГГ-ММ-ДД ЧЧ:ММ:СС", font=("Arial", 8)).pack(pady=2)
        
        # Кнопка для установки полного диапазона
        set_range_button = ttk.Button(time_frame, text="Установить полный диапазон", command=self.set_full_time_range)
        set_range_button.pack(pady=5)
        
        # Кнопка для очистки дат
        clear_dates_button = ttk.Button(time_frame, text="Очистить даты", command=self.clear_dates)
        clear_dates_button.pack(pady=5)
        
        # Кнопка объединения
        merge_button = ttk.Button(left_frame, text="Объединить файлы", command=self.merge_files)
        merge_button.pack(pady=10)
        
        # Создаем центральный фрейм для параметров
        center_frame = ttk.Frame(self.root)
        center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Фрейм для параметров
        self.parameters_frame = ttk.LabelFrame(center_frame, text="Выбор параметров")
        self.parameters_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Создаем канвас и скроллбар для прокрутки параметров
        canvas = tk.Canvas(self.parameters_frame)
        scrollbar = ttk.Scrollbar(self.parameters_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Размещаем элементы
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Кнопки выбора параметров
        select_frame = ttk.Frame(center_frame)
        select_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(select_frame, text="Выбрать все", command=self.select_all_parameters).pack(side=tk.LEFT, padx=5)
        ttk.Button(select_frame, text="Снять все", command=self.deselect_all_parameters).pack(side=tk.LEFT, padx=5)
        
        # Создаем фрейм для правой части (узлы измерения)
        right_frame = ttk.Frame(self.root)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Фрейм для узлов измерения
        self.nodes_frame = ttk.LabelFrame(right_frame, text="Загруженные узлы измерения")
        self.nodes_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Создаем канвас и скроллбар для прокрутки узлов
        nodes_canvas = tk.Canvas(self.nodes_frame)
        nodes_scrollbar = ttk.Scrollbar(self.nodes_frame, orient="vertical", command=nodes_canvas.yview)
        self.nodes_scrollable_frame = ttk.Frame(nodes_canvas)

        self.nodes_scrollable_frame.bind(
            "<Configure>",
            lambda e: nodes_canvas.configure(scrollregion=nodes_canvas.bbox("all"))
        )

        nodes_canvas.create_window((0, 0), window=self.nodes_scrollable_frame, anchor="nw")
        nodes_canvas.configure(yscrollcommand=nodes_scrollbar.set)
        
        # Размещаем элементы
        nodes_canvas.pack(side="left", fill="both", expand=True)
        nodes_scrollbar.pack(side="right", fill="y")
        
        # Кнопки выбора узлов
        nodes_select_frame = ttk.Frame(right_frame)
        nodes_select_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(nodes_select_frame, text="Выбрать все", command=self.select_all_nodes).pack(side=tk.LEFT, padx=5)
        ttk.Button(nodes_select_frame, text="Снять все", command=self.deselect_all_nodes).pack(side=tk.LEFT, padx=5)
        
        # Загружаем параметры из файла правил
        self.load_parameters()
        
        # Авторское право
        copyright_label = ttk.Label(right_frame, text="© Н.А. Галаков", font=("Arial", 8), foreground="gray")
        copyright_label.pack(side=tk.BOTTOM, anchor=tk.E, padx=5, pady=5)
        
    def load_parameters(self):
        """Загружает параметры из файла правил"""
        try:
            # Очищаем старые чекбоксы
            for widget in self.scrollable_frame.winfo_children():
                if isinstance(widget, ttk.Checkbutton):
                    widget.destroy()
            self.parameter_vars.clear()
            
            # Читаем файл с правилами
            rules_file = resource_path("Правила названия столбцов.xlsx")
            if not os.path.exists(rules_file):
                messagebox.showerror("Ошибка", "Файл с правилами названия столбцов не найден!")
                return
                
            # Читаем файл с правилами
            rules_df = pd.read_excel(rules_file, engine='openpyxl')
            
            # Получаем уникальные параметры из 5-го столбца
            parameters = rules_df.iloc[:, 4].unique()
            parameters = [str(p).strip() for p in parameters if pd.notna(p) and str(p).strip()]
            
            # Создаем чекбоксы для каждого параметра
            for param in parameters:
                var = tk.BooleanVar(value=True)
                self.parameter_vars[param] = var
                ttk.Checkbutton(self.scrollable_frame, text=param, variable=var).pack(anchor="w", padx=5, pady=2)
                
        except Exception as e:
            error_message = f"Ошибка при загрузке параметров: {str(e)}"
            print(error_message)
            messagebox.showerror("Ошибка", error_message)
            
    def select_all_parameters(self):
        """Выбирает все параметры"""
        for var in self.parameter_vars.values():
            var.set(True)
            
    def deselect_all_parameters(self):
        """Снимает выбор со всех параметров"""
        for var in self.parameter_vars.values():
            var.set(False)
            
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите Excel файлы",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.files_listbox.insert(tk.END, os.path.basename(file))
                # Обновляем список узлов измерения при добавлении файла
                self.update_measurement_nodes()
                # Обновляем список параметров
                self.load_parameters()
                # Обновляем временной диапазон
                df = self.read_excel_file(file)
                self.update_time_range(df)
                
    def remove_file(self):
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            self.files.pop(index)
            self.files_listbox.delete(index)
            # Обновляем список узлов измерения при удалении файла
            self.update_measurement_nodes()
            # Обновляем список параметров
            self.load_parameters()
            # Сбрасываем временной диапазон
            self.clear_dates()
            
    def select_all_nodes(self):
        """Выбирает все узлы измерения"""
        for var in self.node_vars.values():
            var.set(True)
            
    def deselect_all_nodes(self):
        """Снимает выбор со всех узлов измерения"""
        for var in self.node_vars.values():
            var.set(False)
            
    def update_measurement_nodes(self):
        """Обновляет список узлов измерения на основе загруженных файлов"""
        try:
            # Очищаем старые чекбоксы узлов
            for widget in self.nodes_scrollable_frame.winfo_children():
                if isinstance(widget, ttk.Checkbutton):
                    widget.destroy()
            self.node_vars.clear()
            
            # Читаем файл с правилами
            rules_file = resource_path("Правила названия столбцов.xlsx")
            if not os.path.exists(rules_file):
                messagebox.showerror("Ошибка", "Файл с правилами названия столбцов не найден!")
                return
                
            rules_df = pd.read_excel(rules_file, engine='openpyxl')
            
            # Множество для хранения уникальных узлов измерения
            measurement_nodes = set()
            
            # Для каждого файла находим соответствующие узлы измерения
            for file in self.files:
                filename = unicodedata.normalize('NFC', os.path.basename(file).lower())
                base_filename = unicodedata.normalize('NFC', os.path.splitext(filename)[0])
                logging.info(f"DEBUG: Processing file '{filename}', basename '{base_filename}'")
                
                # Ищем соответствующие правила
                for _, row in rules_df.iterrows():
                    if len(row) >= 4:  # Проверяем наличие столбца с названием узла
                        file_pattern = str(row.iloc[0]).strip().lower()
                        file_pattern = unicodedata.normalize('NFC', file_pattern)
                        node_name = str(row.iloc[3]).strip()  # Предполагаем, что название узла в 4-м столбце
                        
                        # Проверяем, что значения не являются NaN и не пустые
                        if (pd.notna(file_pattern) and pd.notna(node_name) 
                            and file_pattern and node_name 
                            and node_name.lower() != 'nan'):  # Добавляем проверку на 'nan'
                            
                            match = (file_pattern in filename or base_filename in file_pattern)
                            if match:
                                logging.info(f"DEBUG: MATCH FOUND! '{base_filename}' matches pattern '{file_pattern}' -> node '{node_name}'")
                                measurement_nodes.add(node_name)
                            elif '368fq' in filename or '140fq' in filename:
                                # Log near misses
                                if base_filename[:20] == file_pattern[:20]:
                                    logging.info(f"DEBUG: NEAR MISS? file='{base_filename}', pattern='{file_pattern}'")
            
            # Добавляем найденные узлы и создаем чекбоксы
            for node in sorted(measurement_nodes):
                if node and node.lower() != 'nan':
                    # Создаем чекбокс для узла
                    var = tk.BooleanVar(value=True)  # По умолчанию все выбраны
                    self.node_vars[node] = var
                    ttk.Checkbutton(self.nodes_scrollable_frame, text=node, variable=var).pack(anchor="w", padx=5, pady=2)
                
        except Exception as e:
            error_message = f"Ошибка при обновлении списка узлов измерения: {str(e)}"
            print(error_message)
            messagebox.showerror("Ошибка", error_message)
            
    def remove_empty_columns(self, df):
        """Удаляет пустые столбцы из DataFrame"""
        # Удаляем столбцы, где все значения NaN
        empty_cols = df.columns[df.isna().all()].tolist()
        if empty_cols:
            df = df.drop(columns=empty_cols)
        return df

    def read_excel_file(self, file_path):
        """Читает Excel файл с поддержкой обоих форматов .xls и .xlsx"""
        if file_path.endswith('.xlsx'):
            return pd.read_excel(file_path, engine='openpyxl')
        else:  # для .xls файлов
            return pd.read_excel(file_path, engine='xlrd')
            
    def get_rename_rules(self, filename):
        """Получает правила переименования для конкретного файла"""
        try:
            # Читаем файл с правилами
            rules_file = resource_path("Правила названия столбцов.xlsx")
            if not os.path.exists(rules_file):
                messagebox.showerror("Ошибка", "Файл с правилами названия столбцов не найден!")
                return {}
                
            # Читаем файл с правилами
            rules_df = pd.read_excel(rules_file, engine='openpyxl')
            
            # Ищем правила для данного файла
            rename_dict = {}
            filename = unicodedata.normalize('NFC', os.path.basename(filename).lower())  # Получаем имя файла без пути и нормализуем
            
            # Получаем список выбранных параметров
            selected_parameters = [param for param, var in self.parameter_vars.items() if var.get()]
            
            # Перебираем все строки в файле правил
            for _, row in rules_df.iterrows():
                if len(row) >= 5:  # Проверяем наличие нужных столбцов
                    file_pattern = str(row.iloc[0]).strip().lower()
                    file_pattern = unicodedata.normalize('NFC', file_pattern)
                    old_name = str(row.iloc[1]).strip()
                    new_name = str(row.iloc[2]).strip()
                    parameter = str(row.iloc[4]).strip()
                    
                    # Проверяем, что значения не являются NaN и не пустые
                    if (pd.notna(file_pattern) and pd.notna(old_name) and pd.notna(new_name) and pd.notna(parameter)
                        and file_pattern and old_name and new_name and parameter
                        and (file_pattern in filename or os.path.splitext(filename)[0] in file_pattern)
                        and parameter in selected_parameters):
                        rename_dict[old_name] = new_name
                        print(f"Найдено правило для файла '{filename}': '{old_name}' -> '{new_name}' (Параметр: {parameter})")
            
            return rename_dict
            
        except Exception as e:
            error_message = f"Ошибка при чтении правил переименования: {str(e)}"
            print(error_message)
            messagebox.showerror("Ошибка", error_message)
            return {}
            
    def open_range_editor(self):
        """Открывает окно для редактирования диапазонов значений"""
        editor = tk.Toplevel(self.root)
        editor.title("Редактор диапазонов")
        editor.geometry("600x400")
        
        # Создаем фрейм с прокруткой
        canvas = tk.Canvas(editor)
        scrollbar = ttk.Scrollbar(editor, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Загружаем текущие правила
        try:
            rules_file = resource_path("Правила названия столбцов.xlsx")
            if not os.path.exists(rules_file):
                messagebox.showerror("Ошибка", "Файл с правилами не найден")
                return
            
            # Читаем Excel файл
            df = pd.read_excel(rules_file, engine='openpyxl')
            
            entries = {}
            
            # Группируем по имени параметра (NewName - столбец index 2)
            unique_params = df.iloc[:, 2].unique()
            unique_params = [p for p in unique_params if pd.notna(p)]
            unique_params.sort()
            
            for param_name in unique_params:
                # Получаем текущие значения min/max для этого параметра
                param_rows = df[df.iloc[:, 2] == param_name]
                if param_rows.empty:
                    continue
                    
                current_min = param_rows.iloc[0, 5] if len(param_rows.columns) > 5 else ""
                current_max = param_rows.iloc[0, 6] if len(param_rows.columns) > 6 else ""
                
                if pd.isna(current_min): current_min = ""
                if pd.isna(current_max): current_max = ""
                
                # Пропускаем параметры, у которых не заданы ни мин, ни макс значения
                if current_min == "" and current_max == "":
                    continue
                
                # UI для строки
                frame = ttk.Frame(scrollable_frame)
                frame.pack(fill="x", padx=5, pady=2)
                
                ttk.Label(frame, text=param_name, width=30).pack(side="left")
                
                ttk.Label(frame, text="Min:").pack(side="left")
                min_entry = ttk.Entry(frame, width=10)
                min_entry.insert(0, str(current_min))
                min_entry.pack(side="left", padx=5)
                
                ttk.Label(frame, text="Max:").pack(side="left")
                max_entry = ttk.Entry(frame, width=10)
                max_entry.insert(0, str(current_max))
                max_entry.pack(side="left", padx=5)
                
                entries[param_name] = (min_entry, max_entry)
            
            def save_changes():
                try:
                    for param_name, (min_e, max_e) in entries.items():
                        new_min = min_e.get().strip()
                        new_max = max_e.get().strip()
                        
                        try:
                            val_min = float(new_min) if new_min else None
                        except ValueError:
                            val_min = None
                            
                        try:
                            val_max = float(new_max) if new_max else None
                        except ValueError:
                            val_max = None
                            
                        mask = df.iloc[:, 2] == param_name
                        if val_min is not None:
                            df.loc[mask, df.columns[5]] = val_min
                        else:
                            df.loc[mask, df.columns[5]] = None
                            
                        if val_max is not None:
                            df.loc[mask, df.columns[6]] = val_max
                        else:
                            df.loc[mask, df.columns[6]] = None

                    df.to_excel(rules_file, index=False)
                    messagebox.showinfo("Успех", "Диапазоны обновлены успешно")
                    editor.destroy()
                    
                except Exception as e:
                    messagebox.showerror("Ошибка сохранения", str(e))
            
            ttk.Button(editor, text="Сохранить", command=save_changes).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось инициализировать редактор: {e}")

    def update_time_range(self, df):
        """Обновляет диапазон времени на основе данных"""
        try:
            if 'Время' in df.columns:
                time_col = df['Время']
                if pd.api.types.is_datetime64_any_dtype(time_col):
                    start_time = time_col.min()
                    end_time = time_col.max()
                    
                    # Если временной диапазон уже существует, обновляем его
                    if hasattr(self, 'time_range'):
                        current_start, current_end = self.time_range
                        start_time = min(start_time, current_start)
                        end_time = max(end_time, current_end)
                    
                    self.time_range = (start_time, end_time)
                    self.start_date.delete(0, tk.END)
                    self.start_date.insert(0, start_time.strftime('%Y-%m-%d'))
                    self.start_time.delete(0, tk.END)
                    self.start_time.insert(0, start_time.strftime('%H:%M:%S'))
                    
                    self.end_date.delete(0, tk.END)
                    self.end_date.insert(0, end_time.strftime('%Y-%m-%d'))
                    self.end_time.delete(0, tk.END)
                    self.end_time.insert(0, end_time.strftime('%H:%M:%S'))
                    
        except Exception as e:
            print(f"Ошибка при обновлении диапазона времени: {str(e)}")
            
    def set_full_time_range(self):
        """Устанавливает полный диапазон дат из загруженных данных"""
        if hasattr(self, 'time_range'):
            start_time, end_time = self.time_range
            self.start_date.delete(0, tk.END)
            self.start_date.insert(0, start_time.strftime('%Y-%m-%d'))
            self.start_time.delete(0, tk.END)
            self.start_time.insert(0, start_time.strftime('%H:%M:%S'))
            
            self.end_date.delete(0, tk.END)
            self.end_date.insert(0, end_time.strftime('%Y-%m-%d'))
            self.end_time.delete(0, tk.END)
            self.end_time.insert(0, end_time.strftime('%H:%M:%S'))
            
    def clear_dates(self):
        """Очищает поля ввода дат"""
        self.start_date.delete(0, tk.END)
        self.start_time.delete(0, tk.END)
        self.end_date.delete(0, tk.END)
        self.end_time.delete(0, tk.END)
        
    def merge_files(self):
        if not self.files:
            messagebox.showerror("Ошибка", "Пожалуйста, добавьте файлы для объединения")
            return
            
        # Проверяем, выбран ли хотя бы один параметр
        if not any(var.get() for var in self.parameter_vars.values()):
            messagebox.showerror("Ошибка", "Пожалуйста, выберите хотя бы один параметр для объединения")
            return
            
        # Проверяем, выбран ли хотя бы один узел измерения
        if not any(var.get() for var in self.node_vars.values()):
            messagebox.showerror("Ошибка", "Пожалуйста, выберите хотя бы один узел измерения")
            return
            
        try:
            # Проверяем и преобразуем временной интервал
            start_datetime = None
            end_datetime = None
            
            if self.start_date.get() or self.start_time.get():
                try:
                    start_str = f"{self.start_date.get()} {self.start_time.get()}"
                    start_datetime = pd.to_datetime(start_str)
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Неверный формат начальной даты/времени: {str(e)}")
                    return
                    
            if self.end_date.get() or self.end_time.get():
                try:
                    end_str = f"{self.end_date.get()} {self.end_time.get()}"
                    end_datetime = pd.to_datetime(end_str)
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Неверный формат конечной даты/времени: {str(e)}")
                    return
                    
            if start_datetime and end_datetime and start_datetime > end_datetime:
                messagebox.showerror("Ошибка", "Начальная дата/время не может быть позже конечной")
                return
            
            # Читаем файл с правилами для получения соответствия параметров и новых названий
            rules_file = resource_path("Правила названия столбцов.xlsx")
            if not os.path.exists(rules_file):
                messagebox.showerror("Ошибка", "Файл с правилами названия столбцов не найден!")
                return
                
            rules_df = pd.read_excel(rules_file, engine='openpyxl')
            
            # Создаем словарь соответствия параметров и новых названий столбцов
            param_to_columns = {}
            # Создаем словарь соответствия узлов измерения и их столбцов
            node_to_columns = {}
            
            for _, row in rules_df.iterrows():
                if len(row) >= 5:
                    new_name = str(row.iloc[2]).strip()
                    parameter = str(row.iloc[4]).strip()
                    node_name = str(row.iloc[3]).strip()
                    
                    if pd.notna(new_name) and pd.notna(parameter) and pd.notna(node_name) and new_name and parameter and node_name:
                        # Добавляем в словарь параметров
                        if parameter not in param_to_columns:
                            param_to_columns[parameter] = set()
                        param_to_columns[parameter].add(new_name)
                        
                        # Добавляем в словарь узлов
                        if node_name not in node_to_columns:
                            node_to_columns[node_name] = set()
                        node_to_columns[node_name].add(new_name)
            
            # Получаем списки выбранных параметров и узлов
            selected_parameters = [param for param, var in self.parameter_vars.items() if var.get()]
            selected_nodes = [node for node, var in self.node_vars.items() if var.get()]
            
            # Создаем множество допустимых названий столбцов
            allowed_columns = set()
            for param in selected_parameters:
                if param in param_to_columns:
                    allowed_columns.update(param_to_columns[param])
            
            # Фильтруем столбцы по выбранным узлам
            node_allowed_columns = set()
            for node in selected_nodes:
                if node in node_to_columns:
                    node_allowed_columns.update(node_to_columns[node])
            
            # Пересечение множеств для получения только тех столбцов, которые соответствуют и параметрам, и узлам
            allowed_columns = allowed_columns.intersection(node_allowed_columns)
            
            # Чтение всех файлов
            dfs = []
            time_column = None
            
            for i, file in enumerate(self.files):
                print(f"\nОбработка файла: {file}")
                df = self.read_excel_file(file)
                print(f"Столбцы в файле: {list(df.columns)}")
                
                # Удаляем пустые столбцы из каждого файла
                df = self.remove_empty_columns(df)
                
                if i == 0:
                    # В первом файле сохраняем временной столбец
                    time_column = df.iloc[:, 0]  # Предполагаем, что первый столбец - время
                    df = df.iloc[:, 1:]  # Берем все столбцы кроме временного
                else:
                    # В остальных файлах удаляем временной столбец, если он есть
                    if 'Время' in df.columns:
                        df = df.drop(columns=['Время'])
                    # Если первый столбец похож на время (содержит даты или время), удаляем его
                    if pd.api.types.is_datetime64_any_dtype(df.iloc[:, 0]):
                        df = df.iloc[:, 1:]
                
                # Получаем правила переименования для текущего файла
                rename_rules = self.get_rename_rules(file)
                if rename_rules:
                    print(f"Применяем правила переименования для файла {os.path.basename(file)}:")
                    for old_name, new_name in rename_rules.items():
                        print(f"  {old_name} -> {new_name}")
                    # Переименовываем столбцы согласно правилам
                    df = df.rename(columns=rename_rules)
                
                # Оставляем только столбцы, соответствующие выбранным параметрам и узлам
                columns_to_keep = [col for col in df.columns if col in allowed_columns]
                df = df[columns_to_keep]
                
                dfs.append(df)
            
            # Объединение всех датафреймов по столбцам
            merged_df = pd.concat(dfs, axis=1)
            
            # Удаляем дублирующиеся столбцы (например, если один и тот же файл был добавлен дважды)
            # Это критично для избежания ошибок get_loc во время обработки Excel
            merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()].copy()
            
            print("\nСтолбцы после объединения:", list(merged_df.columns))
            
            # Добавляем временной столбец в начало
            merged_df.insert(0, 'Время', time_column)
            
            # Фильтруем по временному интервалу, если он указан
            if start_datetime is not None or end_datetime is not None:
                if start_datetime is not None:
                    merged_df = merged_df[merged_df['Время'] >= start_datetime]
                if end_datetime is not None:
                    merged_df = merged_df[merged_df['Время'] <= end_datetime]
                print(f"Применен фильтр по времени: {start_datetime if start_datetime else 'начало'} - {end_datetime if end_datetime else 'конец'}")
            
            # Удаляем пустые столбцы из объединенного датафрейма
            merged_df = self.remove_empty_columns(merged_df)
            
            # Добавляем столбцы со стрелками перед сохранением
            merged_df, param_to_node = add_arrow_columns(merged_df, rules_file)
            
            # Сортируем по времени перед сохранением, чтобы спарклайны были корректными
            if 'Время' in merged_df.columns:
                 merged_df.sort_values(by='Время', inplace=True)
            
            # Сохранение результата
            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Сохранить объединенный файл"
            )
            
            if output_file:
                try:
                    print(f"Начинаем сохранение результата в файл: {output_file}")
                    
                    # Сбрасываем индекс, чтобы он соответствовал номерам строк в Excel (начиная с 0 -> Row 2)
                    # Это критично для правильной адресации спарклайнов
                    merged_df.reset_index(drop=True, inplace=True)
                    
                    # Преобразуем числовые данные (заменяем запятые на точки и конвертируем в float)
                    # Это необходимо для правильной работы спарклайнов и графиков
                    print("Преобразование данных в числа...")
                    for col in merged_df.columns:
                        if col != 'Время' and not col.endswith('⚠'):
                            try:
                                # Если столбец типа object (строки), пробуем конвертировать
                                if merged_df[col].dtype == 'object':
                                    merged_df[col] = merged_df[col].astype(str).str.replace(',', '.', regex=False)
                                    merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce')
                            except Exception as conv_err:
                                print(f"Не удалось конвертировать столбец {col}: {conv_err}")




                    # Используем xlsxwriter для поддержки спарклайнов
                    with pd.ExcelWriter(output_file, engine='xlsxwriter', datetime_format='yyyy-mm-dd hh:mm:ss') as writer:
                        # Сохраняем основные данные, начиная со 2 строки (индекс 2), чтобы оставить место для заголовков
                        merged_df.to_excel(writer, sheet_name='Данные', index=False, startrow=2, header=False)
                        
                        # Форматируем лист Данные
                        print("Форматируем лист Данные...")
                        format_data_workbook(writer, 'Данные', merged_df, rules_file)
                        
                        # Создаем лист Dashboard
                        print("Создаем лист Dashboard...")
                        create_dashboard_sheet(writer, merged_df, rules_file, node_allowed_columns)
                        
                    print("Данные и Dashboard успешно сохранены")
                    
                    # Закрываем все открытые файлы Excel
                    import gc
                    gc.collect()
                    
                    messagebox.showinfo("Успех", "Файлы успешно объединены, создан Dashboard!")

                except Exception as save_error:
                    error_message = f"Ошибка при сохранении файла: {str(save_error)}"
                    print(error_message)
                    messagebox.showerror("Ошибка", error_message)
                
        except Exception as e:
            error_message = f"Произошла ошибка при объединении файлов: {str(e)}"
            print(error_message)
            messagebox.showerror("Ошибка", error_message)

def create_dashboard_sheet(writer, df, rules_file, allowed_columns):
    """
    Создает лист Dashboard с Timeline Heatmap и Sparklines.
    """
    try:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Dashboard')
        
        # Базовые шрифты
        font_name = 'Arial'
        
        # Стили
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1,
            'bg_color': '#D9D9D9',
            'font_name': font_name,
            'font_size': 10
        })
        
        node_format = workbook.add_format({
            'bold': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1,
            'font_name': font_name,
            'font_size': 10
        })
        
        # Для пустой ячейки (A2)
        empty_corner_format = workbook.add_format({
            'border': 1,
            'bg_color': '#FFFFFF'
        })
        
        # Для "Итого:" (правое выравнивание)
        itogo_label_format = workbook.add_format({
            'bold': True,
            'valign': 'vcenter',
            'align': 'right',
            'border': 1,
            'bg_color': '#FFFFFF',
            'font_name': font_name,
            'font_size': 14
        })
        
        # Цвета для статусов (с форматом чисел)
        num_format = '# ##0.00'
        
        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1, 'num_format': num_format, 'valign': 'vcenter', 'align': 'center', 'bold': True, 'font_size': 10, 'font_name': font_name, 'font_color': '#000000'})
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1, 'num_format': num_format, 'valign': 'vcenter', 'align': 'center', 'bold': True, 'font_size': 10, 'font_name': font_name, 'font_color': '#000000'})
        grey_format = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'num_format': num_format, 'valign': 'vcenter', 'align': 'center', 'bold': True, 'font_size': 10, 'font_name': font_name, 'font_color': '#000000'})
        yellow_format = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'num_format': num_format, 'valign': 'vcenter', 'align': 'center', 'bold': True, 'font_size': 10, 'font_name': font_name, 'font_color': '#000000'})
        
        # 1. Подготовка данных
        # Определяем столбцы расхода для каждого узла
        # Считываем правила снова, чтобы найти Qmin/Qmax для каждого параметра
        rules_df = pd.read_excel(rules_file, engine='openpyxl')
        
        # Структура: {NodeName: {FlowCol: col_name, Qmin: val, Qmax: val, Units: str}}
        node_config = {}
        
        # Находим столбцы расхода
        for col in df.columns:
            if col not in allowed_columns:
                continue
                
            # Ищем правило для этого столбца (по NewName)
            rule_row = rules_df[rules_df.iloc[:, 2] == col]
            if not rule_row.empty:
                param_name = str(rule_row.iloc[0, 4]).strip()
                node_name = str(rule_row.iloc[0, 3]).strip()
                
                # Простейшая эвристика для определения "расхода"
                if "расход" in param_name.lower():
                    qmin = rule_row.iloc[0, 5] if len(rule_row.columns) > 5 else None
                    qmax = rule_row.iloc[0, 6] if len(rule_row.columns) > 6 else None
                    units = str(rule_row.iloc[0, 7]).strip() if len(rule_row.columns) > 7 and pd.notna(rule_row.iloc[0, 7]) else ""

                    if not units:
                        if "расход" in param_name.lower(): units = "тыс. м3/ч"
                    
                    try:
                        qmin = float(qmin) if pd.notna(qmin) else 0
                        qmax = float(qmax) if pd.notna(qmax) else float('inf')
                    except:
                        qmin, qmax = 0, float('inf')
                        
                    node_config[node_name] = {
                        'col': col,
                        'qmin': qmin,
                        'qmax': qmax,
                        'units': units,
                        'col_idx': df.columns.get_loc(col) # Индекс в dataframe (0-based)
                    }

        if not node_config:
            worksheet.write(0, 0, "Не найдены параметры расхода для построения Dashboard")
            return

        # Группировка по датам (дни)
        if 'Время' not in df.columns:
             worksheet.write(0, 0, "Ошибка: нет столбца Время")
             return

        df['Date'] = pd.to_datetime(df['Время']).dt.date
        unique_dates = sorted(df['Date'].dropna().unique())
        unique_nodes = sorted(node_config.keys())
        
        num_days = len(unique_dates)
        
        # 2. Заголовки
        worksheet.write(0, 0, "Узел (позиция)", header_format)
        worksheet.set_column(0, 0, 18) # Ширина первого столбца
        
        worksheet.write(0, 1, "Допустимый диапазон", header_format)
        worksheet.set_column(1, 1, 35) # Ширина второго столбца
        
        # Новый столбец статистики с динамическим заголовком
        worksheet.write(0, 2, f"Выход за диапазон\nза период {num_days} суток", header_format)
        worksheet.set_column(2, 2, 22)
        
        worksheet.freeze_panes(2, 3)   # Закрепить заголовки (2 строки) и первые три столбца
        
        for j, date in enumerate(unique_dates):
            worksheet.write(0, j + 3, date.strftime('%d.%m.%Y'), header_format)
            worksheet.set_column(j + 3, j + 3, 22) # Ширина столбцов с датами
            
        # 3. Заполнение матрицы
        # Словарь {date: [start_row, end_row]} (0-based indices in DF)
        date_ranges = {}
        # Предполагаем, что df отсортирован по времени
        for date in unique_dates:
            # Находим индексы строк для этой даты
            mask = df['Date'] == date
            indices = df.index[mask].tolist()
            if indices:
                start = indices[0]
                end = indices[-1]
                date_ranges[date] = (start, end)
        
        # Словари для подсчета итогов
        total_days_sum = 0
        total_hours_sum = 0
        daily_sums = {date: 0.0 for date in unique_dates}
        
        # Оформление строки Итого (строка 1)
        worksheet.write(1, 0, "", empty_corner_format)
        worksheet.write(1, 1, "Итого:", itogo_label_format)
        
        # Формат для итогов с единицами измерения (белый фон)
        total_data_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#FFFFFF', # Белый фон
            'border': 1,
            'font_name': font_name,
            'font_size': 12,
            'num_format': '# ##0" (тыс. м3)"'
        })
        
        for i, node in enumerate(unique_nodes):
            row_idx = i + 2 # Смещаем на 2 (Заголовок + Итого)
            worksheet.set_row(row_idx, 40) # Увеличиваем высоту строки для наглядности графика
            
            config = node_config[node]
            qmin = config['qmin']
            qmax = config['qmax']
            units = config['units']
            
            # Формируем раздельные подписи
            qmin_str = f"{qmin:g}"
            if qmax == float('inf'):
                qmax_str = "∞"
            else:
                qmax_str = f"{qmax:g}"
            
            range_label = f"({qmin_str} ... {qmax_str} {units})"
            
            worksheet.write(row_idx, 0, node, node_format)
            worksheet.write(row_idx, 1, range_label, node_format)
            
            col_name = config['col']
            col_idx_in_data = config['col_idx'] # 0-based column index in Data sheet
            
            col_letter = get_column_letter(col_idx_in_data + 1) # Excel 1-based letter
            
            # Переменные для статистики по узлу
            total_violation_days = 0
            total_violation_hours = 0
            
            for j, date in enumerate(unique_dates):
                col_idx = j + 3 # Смещаем на 3 столбца (Узел, Диапазон, Статистика)
                
                if date in date_ranges:
                    start_row, end_row = date_ranges[date]
                    # Convert to Excel row numbers (1-based)
                    # Sheet 'Данные': Row 1 is header. Data starts Row 2.
                    # DF index 0 -> Excel Row 2.
                    excel_start = start_row + 2
                    excel_end = end_row + 2
                    
                    # Избегаем ошибок с дублирующимися колонками, используя .item() или явно [0] если get_loc вернул массив (хоть мы их и удалили)
                    loc = config['col_idx']
                    idx = int(loc[0]) if isinstance(loc, (list, pd.core.arrays.boolean.BooleanArray)) or hasattr(loc, '__iter__') else int(loc)
                    col_letter = get_column_letter(idx + 1)
                    
                    data_range = f"'Данные'!{col_letter}{excel_start}:{col_letter}{excel_end}"
                    
                    # Анализ данных для определния цвета фона (статус) и суммы
                    values = df.iloc[start_row:end_row+1, idx]
                    vals = pd.to_numeric(values, errors='coerce').fillna(0)
                    
                    day_sum = vals.sum()
                    daily_sums[date] += day_sum # Суммируем для итога
                    
                    status_format = grey_format # Default
                    is_zero = False
                    
                    # Статистика нарушений за день
                    day_violation_hours = 0
                    
                    if vals.max() == 0 and vals.min() == 0:
                         status_format = grey_format
                         is_zero = True
                    else:
                        has_violation = False
                        
                        # Check min (exclude 0)
                        # Считаем количество часов с нарушением мин
                        min_violations = (vals < qmin) & (vals != 0)
                        
                        # Check max
                        max_violations = vals > qmax
                        
                        # Общее количество часов с нарушениями
                        violations_mask = min_violations | max_violations
                        day_violation_hours = violations_mask.sum()
                        
                        if day_violation_hours > 0:
                            has_violation = True
                            
                        # Присвоение цвета
                        if day_sum <= 12:
                            status_format = yellow_format
                            if has_violation:
                                total_violation_days += 1
                        elif has_violation:
                            status_format = red_format
                            total_violation_days += 1
                        else:
                            status_format = green_format
                            
                    total_violation_hours += day_violation_hours
                    
                    # Пишем сумму в ячейку с форматом фона
                    worksheet.write_number(row_idx, col_idx, day_sum, status_format)
                    
                    # Рисуем график только если не ноль
                    if not is_zero:
                        spark_color = '#595959' 
                        if status_format == red_format:
                             spark_color = '#A54040' # Еще более бледный темно-красный
                        elif status_format == green_format:
                             spark_color = '#407040' # Еще более бледный темно-зеленый
                        elif status_format == yellow_format:
                             spark_color = '#B38600' # Темно-желтый для графика
                             
                        options = {
                            'range': data_range,
                            'type': 'line',
                            'markers': False,
                            'weight': 1.0, # Тоньше линия
                            'series_color': spark_color,
                            'high_point': False,
                            'low_point': False,
                        }
                        worksheet.add_sparkline(row_idx, col_idx, options)
                    
                else:
                    worksheet.write(row_idx, col_idx, "Н/Д", grey_format)
            
            # Накапливаем итоговую статистику
            total_days_sum += total_violation_days
            total_hours_sum += total_violation_hours
            
            # Записываем статистику в столбец 2 (индекс 2)
            stats_text = f"{total_violation_days} сут.; {int(total_violation_hours)} ч."
            stats_fmt = green_format
            if total_violation_hours > 0:
                stats_fmt = red_format
            
            worksheet.write(row_idx, 2, stats_text, stats_fmt)
            
        # Записываем итоги в строку 1
        # Итог по статистике
        total_stats_text = f"{total_days_sum} сут.; {int(total_hours_sum)} ч."
        total_stats_fmt = workbook.add_format({
            'bold': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1,
            'bg_color': '#FFFFFF',
            'font_name': font_name,
            'font_size': 14,
            'font_color': '#FF0000'
        })
        worksheet.write(1, 2, total_stats_text, total_stats_fmt)
        
        # Итоги по датам
        for j, date in enumerate(unique_dates):
            col_idx = j + 3
            val = daily_sums.get(date, 0)
            worksheet.write_number(1, col_idx, val, total_data_format)
            
        # Делаем лист активным при открытии
        worksheet.activate()

    except Exception as e:
        print(f"Ошибка при создании Dashboard: {e}")
        import traceback
        traceback.print_exc()


def setup_logging():
    """Configures logging to a file in the user's home directory."""
    log_dir = os.path.join(os.path.expanduser("~"), ".analytics_ui")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "app.log")
    
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    return log_file

def main():
    try:
        log_file = setup_logging()
        print(f"Logging to: {log_file}")
        print("Starting application...")
        
        root = tk.Tk()
        app = ExcelMerger(root)
        print("Entering main loop...")
        root.mainloop()
    except Exception as e:
        error_msg = f"Critical error: {e}"
        print(error_msg)
        logging.critical(error_msg, exc_info=True)
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...") # Keep window open if run from double-click

if __name__ == "__main__":
    main() 