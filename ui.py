import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from file_operations import export_to_format
from file_operations import evaluate_formulas




class ExcelToCsvApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Tool")
        self.root.geometry("850x600")
        self.current_df = pd.DataFrame()  # Исходный DataFrame
        self.filtered_df = pd.DataFrame()  # Отфильтрованный DataFrame

        # Создаём вкладки
        self.create_tabs()


        # Переменные для текущих данных
        self.current_df = None
        self.filtered_df = None



    def create_tabs(self):
        """
        Создаёт вкладки для интерфейса.
        """
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        # Вкладка для работы с файлами
        self.file_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.file_tab, text="Работа с файлами")

        # Вкладка для предварительного просмотра
        self.preview_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.preview_tab, text="Предварительный просмотр")

        # Создание интерфейса на вкладках
        self.create_file_tab_widgets()
        self.create_preview_tab_widgets()

    def sort_column(self, col, reverse):
        """
        Сортирует данные в таблице по выбранной колонке.
        """
        try:
            data = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
            try:
                data.sort(key=lambda t: float(t[0].replace(',', '.')), reverse=reverse)
            except ValueError:
                data.sort(key=lambda t: t[0], reverse=reverse)

            for index, (_, k) in enumerate(data):
                self.tree.move(k, '', index)

            self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))
        except Exception as e:
            messagebox.showerror("Ошибка сортировки", f"Не удалось отсортировать данные: {e}")

    def create_file_tab_widgets(self):
        """load_csv_file
        Создаёт виджеты для вкладки "Работа с файлами".
        """
        label = tk.Label(self.file_tab, text="Выберите файл Excel:")
        label.pack(pady=5)

        self.file_entry = tk.Entry(self.file_tab, width=50, state="readonly")
        self.file_entry.pack(pady=5)

        browse_button = tk.Button(self.file_tab, text="Обзор", command=self.select_file)
        browse_button.pack(pady=5)

        self.sheet_label = tk.Label(self.file_tab, text="Выберите лист:")
        self.sheet_label.pack(pady=5)

        self.sheet_combo = ttk.Combobox(self.file_tab, state="readonly")
        self.sheet_combo.pack(pady=5)

        preview_button = tk.Button(self.file_tab, text="Показать данные", command=self.preview_data)
        preview_button.pack(pady=10)

        self.format_label = tk.Label(self.file_tab, text="Выберите формат для сохранения:")
        self.format_label.pack(pady=5)

        self.format_combo = ttk.Combobox(
            self.file_tab, state="readonly", values=["CSV", "JSON", "XML", "Markdown", "YAML", "SQL", "HTML"]
        )
        self.format_combo.pack(pady=5)
        self.format_combo.current(0)

        save_button = tk.Button(self.file_tab, text="Сохранить", command=self.export_file)
        save_button.pack(pady=10)

        # Добавляем кнопку для конвертации CSV в Excel
        convert_button = tk.Button(self.file_tab, text="Конвертировать CSV в Excel", command=self.export_csv_to_excel)
        convert_button.pack(pady=5)

    def create_preview_tab_widgets(self):
        """
        Создаёт виджеты для вкладки "Предварительный просмотр".
        """
        self.tree = ttk.Treeview(self.preview_tab)
        self.tree.pack(pady=10, fill=tk.BOTH, expand=True)

        # Добавление кнопки для сохранения отфильтрованных данных в Excel
        save_excel_button = tk.Button(self.preview_tab, text="Сохранить в Excel", command=self.save_filtered_to_excel)
        save_excel_button.pack(pady=10)

    def select_file(self):
        """
        Открывает диалог для выбора файла и загружает данные в приложение.
        """
        file_path = filedialog.askopenfilename(
            title="Выберите файл Excel или CSV",
            filetypes=[("Excel и CSV файлы", "*.xlsx *.xls *.csv"), ("Все файлы", "*.*")]
        )
        if file_path:
            self.file_entry.config(state="normal")
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.file_entry.config(state="readonly")
            self.excel_file = file_path

            if file_path.endswith(('.xlsx', '.xls')):
                self.load_sheets(file_path)
            elif file_path.endswith('.csv'):
                self.load_csv(file_path)
            else:
                messagebox.showerror("Ошибка", "Неподдерживаемый формат файла.")
        else:
            messagebox.showerror("Ошибка", "Файл не выбран.")

    def load_csv(self, file_path):
        """
        Загружает CSV-файл и отображает его в предварительном просмотре.
        """
        try:
            # Загружаем CSV в DataFrame
            self.current_df = pd.read_csv(file_path)

            # Сбрасываем фильтры и связанные данные
            self.filtered_df = None  # Обнуляем отфильтрованные данные
            self.sheet_combo['values'] = ["CSV"]  # Устанавливаем фиктивное значение для выбора листа
            self.sheet_combo.current(0)  # Выбираем первый (и единственный) элемент
            self.update_table(self.current_df)  # Обновляем таблицу для отображения данных

            # Создаём интерфейс для фильтрации (если его нет)
            self.create_filter_widgets()

            # Показываем сообщение об успешной загрузке
            messagebox.showinfo("Успех", f"Данные из CSV успешно загружены:\n{file_path}")
        except pd.errors.EmptyDataError:
            messagebox.showerror("Ошибка", "Файл пуст или содержит некорректные данные.")
        except pd.errors.ParserError:
            messagebox.showerror("Ошибка", "Ошибка парсинга файла. Проверьте формат CSV.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить CSV: {e}")

    def load_sheets(self, file_path):
        try:
            excel_file = pd.ExcelFile(file_path)
            self.sheet_combo['values'] = excel_file.sheet_names
            self.sheet_combo.current(0)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить листы: {e}")

    def update_table(self, data):
        """
        Обновляет таблицу в интерфейсе для предварительного просмотра данных.
        """
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(data.columns)
        self.tree["show"] = "headings"

        for col in data.columns:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_column(_col, False))
            self.tree.column(col, width=100)

        for _, row in data.iterrows():
            self.tree.insert("", "end", values=list(row))

    def preview_data(self):
        """
        Отображает предварительный просмотр данных с возможностью фильтрации.
        """
        file_path = self.file_entry.get().strip()
        if not file_path:
            messagebox.showerror("Ошибка", "Файл не выбран.")
            return

        selected_sheet = self.sheet_combo.get().strip() if self.sheet_combo['values'] else None

        try:
            # Загружаем данные из Excel
            if file_path.endswith(('.xlsx', '.xls')):
                self.current_df = pd.read_excel(file_path, sheet_name=selected_sheet)
            else:
                messagebox.showerror("Ошибка", "Неподдерживаемый формат файла.")
                return

            # Применяем формулы
            self.current_df = evaluate_formulas(self.current_df)

            # Обновляем таблицу для предварительного просмотра
            self.update_table(self.current_df)

            # Создаём интерфейс для фильтрации
            self.create_filter_widgets()

            messagebox.showinfo("Успех", "Данные успешно загружены.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")

    def create_filter_widgets(self):
        """
        Создаёт интерфейс для управления фильтрами, включая выпадающие списки с уникальными значениями.
        """
        # Удаляем старый фильтр-фрейм, если он уже существует
        if hasattr(self, 'filter_frame'):
            self.filter_frame.destroy()
        if hasattr(self, 'filters_container'):
            self.filters_container.destroy()

        # Создаём контейнер для управления фильтрами
        self.filter_frame = tk.Frame(self.preview_tab)
        self.filter_frame.pack(pady=10, fill=tk.X)

        self.filters = []  # Список для хранения активных фильтров

        # Кнопки управления фильтрами
        add_filter_button = tk.Button(self.filter_frame, text="Добавить фильтр", command=self.add_filter)
        add_filter_button.pack(side=tk.LEFT, padx=5)

        apply_filters_button = tk.Button(self.filter_frame, text="Применить фильтры", command=self.apply_filters)
        apply_filters_button.pack(side=tk.LEFT, padx=5)

        save_button = tk.Button(self.filter_frame, text="Сохранить в CSV", command=self.save_filtered_data)
        save_button.pack(side=tk.LEFT, padx=5)

        # Контейнер для отображения активных фильтров
        self.filters_container = tk.Frame(self.preview_tab)
        self.filters_container.pack(pady=10, fill=tk.BOTH, expand=True)

        # Автоматическое добавление первого фильтра
        self.add_filter()

    def add_filter(self):
        """
        Добавляет новый фильтр в интерфейс.
        """
        filter_row = tk.Frame(self.filters_container)
        filter_row.pack(fill=tk.X, pady=5)

        # Выпадающий список для выбора столбца
        column_label = tk.Label(filter_row, text="Столбец:")
        column_label.pack(side=tk.LEFT, padx=5)

        column_combobox = ttk.Combobox(
            filter_row,
            values=list(self.current_df.columns),  # Уникальные значения для столбцов
            state="readonly",
            width=20
        )
        column_combobox.pack(side=tk.LEFT, padx=5)

        # Выпадающий список для выбора значения
        value_label = tk.Label(filter_row, text="Значение:")
        value_label.pack(side=tk.LEFT, padx=5)

        value_combobox = ttk.Combobox(filter_row, state="readonly", width=45)
        value_combobox.pack(side=tk.LEFT, padx=5)

        def update_values(event):
            column = column_combobox.get()
            if column:
                unique_values = self.current_df[column].dropna().unique()  # Убираем NaN
                value_combobox['values'] = list(unique_values)

        column_combobox.bind("<<ComboboxSelected>>", update_values)

        # Кнопка для удаления фильтра
        delete_button = tk.Button(filter_row, text="Удалить", command=lambda: self.remove_filter(filter_row))
        delete_button.pack(side=tk.LEFT, padx=5)

        # Сохраняем виджеты фильтра
        self.filters.append((column_combobox, value_combobox, filter_row))

    def remove_filter(self, filter_row):
        """
        Удаляет фильтр из интерфейса.
        """
        for i, (column_combobox, value_combobox, row) in enumerate(self.filters):
            if row == filter_row:
                self.filters.pop(i)
                row.destroy()
                break
        filter_row.destroy()

    def apply_filters(self):
        """
        Применяет все активные фильтры к данным, поддерживая множественные фильтры для одного столбца.
        """
        if self.current_df.empty:
            messagebox.showwarning("Предупреждение", "Данные отсутствуют для фильтрации.")
            return

        filtered_df = self.current_df.copy()  # Создаём копию исходного DataFrame
        column_filters = {}  # Словарь для хранения фильтров по столбцам

        try:
            # Собираем все фильтры в словарь по столбцам
            for i, (column_combo, value_combo, row) in enumerate(self.filters):
                column = column_combo.get().strip()  # Получаем выбранную колонку
                value = value_combo.get().strip()  # Получаем выбранное значение из выпадающего списка

                if column and value:
                    if column not in column_filters:
                        column_filters[column] = set()  # Создаём множество для фильтров по столбцу
                    column_filters[column].add(value)  # Добавляем значение в множество
                else:
                    messagebox.showwarning(
                        "Предупреждение",
                        f"Фильтр {i + 1}: Колонка или значение не указаны. Пропускаем."
                    )

            # Применяем фильтры ко всем колонкам
            for column, values in column_filters.items():
                # Отфильтровываем строки, где значения столбца совпадают с одним из указанных
                filtered_df = filtered_df[filtered_df[column].astype(str).isin(values)]

            # Обновляем таблицу Treeview и сохраняем результат
            self.update_table(filtered_df)
            self.filtered_df = filtered_df  # Сохраняем отфильтрованный DataFrame для дальнейшего использования

            # Уведомляем об успешном применении фильтров
            messagebox.showinfo("Успех", f"Фильтры применены. Найдено {len(filtered_df)} строк.")
        except KeyError as e:
            # Ошибка доступа к колонке (например, если она была удалена)
            messagebox.showerror("Ошибка", f"Указана недействительная колонка: {e}")
        except Exception as e:
            # Прочие ошибки
            messagebox.showerror("Ошибка", f"Не удалось применить фильтры: {e}")
    def save_filtered_data(self):
        """
        Сохраняет отфильтрованные данные в CSV.
        """
        if not hasattr(self, "filtered_df") or self.filtered_df is None:
            messagebox.showerror("Ошибка", "Нет данных для сохранения. Сначала примените фильтры.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Сохранить отфильтрованные данные",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )

        if save_path:
            try:
                self.filtered_df.to_csv(save_path, index=False, encoding="utf-8-sig")
                messagebox.showinfo("Успех", "Файл успешно сохранён!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")

    def save_filtered_to_excel(self):
        """
        Сохраняет текущие данные (отфильтрованные или исходные) в Excel-файл.
        """
        try:
            # Используем отфильтрованный DataFrame, если фильтры применены, иначе исходный
            df_to_save = self.filtered_df if self.filtered_df is not None else self.current_df

            # Проверяем, что данные существуют и не пусты
            if df_to_save is None or df_to_save.empty:
                messagebox.showerror("Ошибка", "Нет данных для сохранения.")
                return

            # Открываем диалог для выбора пути сохранения
            file_path = filedialog.asksaveasfilename(
                title="Сохранить как Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")]
            )
            if not file_path:  # Если пользователь отменил выбор пути
                return

            # Сохраняем данные в Excel
            df_to_save.to_excel(file_path, index=False, engine="openpyxl")
            messagebox.showinfo("Успех", f"Данные успешно сохранены в файл:\n{file_path}")
        except Exception as e:
            # Выводим сообщение об ошибке
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")

    def export_csv_to_excel(self):
        """
        Конвертирует CSV в Excel.
        """
        file_path = filedialog.askopenfilename(
            title="Выберите CSV файл",
            filetypes=[("CSV Files", "*.csv"), ("Все файлы", "*.*")]
        )
        if not file_path:
            messagebox.showerror("Ошибка", "Файл не выбран.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Сохранить как Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            # Чтение CSV и сохранение в Excel
            df = pd.read_csv(file_path)
            df.to_excel(save_path, index=False, engine="openpyxl")
            messagebox.showinfo("Успех", "Файл успешно сохранён как Excel!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось конвертировать файл: {e}")

    def apply_filter(self):
        """
        Применяет фильтр к данным на основе выбранной колонки и введённого значения.
        """
        column = self.filter_column_combo.get()
        value = self.filter_value_entry.get().strip()

        if not column or not value:
            messagebox.showerror("Ошибка", "Выберите колонку и введите значение для фильтра.")
            return

        try:
            # Фильтруем данные
            filtered_df = self.current_df[self.current_df[column].astype(str).str.contains(value, case=False, na=False)]
            self.update_table(filtered_df)
            messagebox.showinfo("Успех", f"Фильтр применён. Найдено {len(filtered_df)} строк.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось применить фильтр: {e}")

    def export_file(self):
        """
        Экспорт данных в выбранный формат.
        """
        if not hasattr(self, "current_df") or self.current_df is None:
            messagebox.showerror("Ошибка", "Сначала загрузите и просмотрите данные.")
            return

        selected_format = self.format_combo.get()
        save_path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=f".{selected_format.lower()}",
            filetypes=[(f"{selected_format.upper()} Files", f"*.{selected_format.lower()}")],
        )
        if save_path:
            try:
                export_to_format(self.current_df, selected_format, save_path)
                messagebox.showinfo("Успех", f"Файл успешно сохранён в формате {selected_format}!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")



if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToCsvApp(root)
    root.mainloop()