import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QComboBox, QTreeView, QFileDialog, QMessageBox, QTabWidget, QFrame, QAbstractItemView
)
from PyQt5.QtCore import Qt, QAbstractTableModel, QVariant

from file_operations import export_to_format


class PandasModel(QAbstractTableModel):
    """
    Модель для отображения DataFrame в QTableView.
    """
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])
        return QVariant()

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return str(self._data.columns[section])
        return QVariant()


class ExcelToCsvApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.save_button = None
        self.convert_button = None
        self.browse_button = None
        self.add_filter_button = None
        self.current_df = pd.DataFrame()  # Исходный DataFrame
        self.filtered_df = pd.DataFrame()  # Отфильтрованный DataFrame
        self.init_ui()

    def init_ui(self):
        """
        Инициализация интерфейса.
        """
        self.setWindowTitle("Excel Data Tool")
        self.setGeometry(100, 100, 850, 600)

        layout = QVBoxLayout()

        # Создаем вкладки
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Вкладка для работы с файлами
        self.file_tab = QWidget()
        self.tabs.addTab(self.file_tab, "Работа с файлами")

        # Вкладка для предварительного просмотра
        self.preview_tab = QWidget()
        self.tabs.addTab(self.preview_tab, "Предварительный просмотр")

        # Создаем виджеты для вкладок
        self.create_file_tab_widgets()
        self.create_preview_tab_widgets()

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def create_file_tab_widgets(self):
        """
        Создает виджеты для вкладки "Работа с файлами".
        """
        layout = QVBoxLayout()

        # Выбор файла
        self.file_label = QLabel("Выберите файл Excel:")
        layout.addWidget(self.file_label)



        self.file_entry = QLineEdit()
        self.file_entry.setReadOnly(True)
        layout.addWidget(self.file_entry)

        self.browse_button = QPushButton("Обзор")
        self.browse_button.setFixedSize(150, 25)
        layout.addWidget(self.browse_button, alignment=Qt.AlignCenter)
        self.browse_button.clicked.connect(self.select_file)


        # Выбор листа
        self.sheet_label = QLabel("Выберите лист:")
        layout.addWidget(self.sheet_label)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setEnabled(False)
        layout.addWidget(self.sheet_combo)

        # Кнопка для предварительного просмотра
        self.preview_button = QPushButton("Показать данные")
        layout.addWidget(self.preview_button, alignment=Qt.AlignCenter)
        self.preview_button.clicked.connect(self.select_file)
        self.preview_button.clicked.connect(self.preview_data)
        # layout.addWidget(self.preview_button)

        # Выбор формата для сохранения
        self.format_label = QLabel("Выберите формат для сохранения:")
        layout.addWidget(self.format_label)

        self.format_combo = QComboBox()
        self.format_combo.addItems(["CSV", "JSON", "XML", "Markdown", "YAML", "SQL", "HTML"])
        self.format_combo.setFixedSize(150, 30)
        layout.addWidget(self.format_combo)


        # Кнопка для сохранения
        self.save_button = QPushButton("Сохранить")
        self.save_button.clicked.connect(self.export_file)
        layout.addWidget(self.save_button, alignment=Qt.AlignCenter)
        self.save_button.clicked.connect(self.select_file)
        # layout.addWidget(self.save_button)

        # Кнопка для конвертации CSV в Excel
        self.convert_button = QPushButton("Конвертировать CSV в Excel")
        layout.addWidget(self.convert_button, alignment=Qt.AlignCenter)
        self.convert_button.clicked.connect(self.select_file)
        self.convert_button.clicked.connect(self.export_csv_to_excel)
        # layout.addWidget(self.convert_button)

        self.file_tab.setLayout(layout)

    def create_preview_tab_widgets(self):
        """
        Создает виджеты для вкладки "Предварительный просмотр".
        """
        layout = QVBoxLayout()

        # Таблица для отображения данных
        self.tree = QTreeView()
        self.tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(self.tree)

        # Кнопка для сохранения отфильтрованных данных в Excel
        self.save_excel_button = QPushButton("Сохранить в Excel")
        self.save_excel_button.setFixedSize(150, 25)
        layout.addWidget(self.save_excel_button, alignment=Qt.AlignCenter)
        self.save_excel_button.clicked.connect(self.save_filtered_to_excel)
        # layout.addWidget(self.save_excel_button)

        # Кнопка для сохранения отфильтрованных данных в CSV
        self.save_csv_button = QPushButton("Сохранить в CSV")
        self.save_csv_button.setFixedSize(150, 25)
        layout.addWidget(self.save_csv_button, alignment=Qt.AlignCenter)
        self.save_csv_button.clicked.connect(self.save_filtered_to_csv)
        # layout.addWidget(self.save_csv_button)

        # Контейнер для фильтров
        self.filter_frame = QFrame()
        self.filter_frame.setLayout(QVBoxLayout())
        layout.addWidget(self.filter_frame)

        # Кнопка для добавления фильтра
        self.add_filter_button = QPushButton("Добавить фильтр")
        self.add_filter_button.setFixedSize(150, 25)
        layout.addWidget(self.add_filter_button, alignment=Qt.AlignCenter)
        self.add_filter_button.clicked.connect(self.add_filter)
        self.filter_frame.layout().addWidget(self.add_filter_button)

        # Кнопка для применения фильтров
        self.apply_filters_button = QPushButton("Применить фильтры")
        self.apply_filters_button.setFixedSize(150, 25)
        layout.addWidget(self.apply_filters_button, alignment=Qt.AlignCenter)
        self.apply_filters_button.clicked.connect(self.apply_filters)
        self.filter_frame.layout().addWidget(self.apply_filters_button)

        self.preview_tab.setLayout(layout)

    def select_file(self):
        """
        Открывает диалог для выбора файла и загружает данные в приложение.
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel или CSV", "", "Excel и CSV файлы (*.xlsx *.xls *.csv);;Все файлы (*.*)"
        )
        if file_path:
            self.file_entry.setText(file_path)
            self.excel_file = file_path

            if file_path.endswith(('.xlsx', '.xls')):
                self.load_sheets(file_path)
            elif file_path.endswith('.csv'):
                self.load_csv(file_path)
            else:
                QMessageBox.critical(self, "Ошибка", "Неподдерживаемый формат файла.")
        else:
            QMessageBox.critical(self, "Ошибка", "Файл не выбран.")

    def load_csv(self, file_path):
        """
        Загружает CSV-файл и отображает его в предварительном просмотре.
        """
        try:
            self.current_df = pd.read_csv(file_path)
            self.filtered_df = None
            self.sheet_combo.clear()
            self.sheet_combo.addItem("CSV")
            self.sheet_combo.setCurrentIndex(0)
            self.update_table(self.current_df)
            QMessageBox.information(self, "Успех", f"Данные из CSV успешно загружены:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить CSV: {e}")

    def load_sheets(self, file_path):
        """
        Загружает листы из Excel-файла.
        """
        try:
            excel_file = pd.ExcelFile(file_path)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(excel_file.sheet_names)
            self.sheet_combo.setCurrentIndex(0)
            self.sheet_combo.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить листы: {e}")

    def update_table(self, data):
        """
        Обновляет таблицу в интерфейсе для предварительного просмотра данных.
        """
        model = PandasModel(data)
        self.tree.setModel(model)

    def preview_data(self):
        """
        Отображает предварительный просмотр данных.
        """
        file_path = self.file_entry.text().strip()
        if not file_path:
            QMessageBox.critical(self, "Ошибка", "Файл не выбран.")
            return

        selected_sheet = self.sheet_combo.currentText().strip() if self.sheet_combo.count() > 0 else None

        try:
            if file_path.endswith(('.xlsx', '.xls')):
                self.current_df = pd.read_excel(file_path, sheet_name=selected_sheet)
            else:
                QMessageBox.critical(self, "Ошибка", "Неподдерживаемый формат файла.")
                return

            self.update_table(self.current_df)
            QMessageBox.information(self, "Успех", "Данные успешно загружены.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить данные: {e}")

    def add_filter(self):
        """
        Добавляет новый фильтр в интерфейс.
        """
        if self.current_df.empty:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите данные!")
            return

        filter_row = QFrame()
        filter_row.setLayout(QHBoxLayout())

        # Создаем выпадающий список для столбцов
        column_label = QLabel("Столбец:")
        filter_row.layout().addWidget(column_label)

        column_combobox = QComboBox()
        column_combobox.addItem("")  # Добавляем пустой элемент
        column_combobox.addItems(self.current_df.columns)
        column_combobox.setCurrentIndex(-1)  # Устанавливаем пустое значение по умолчанию
        filter_row.layout().addWidget(column_combobox)

        # Создаем выпадающий список для значений
        value_label = QLabel("Значение:")
        filter_row.layout().addWidget(value_label)

        value_combobox = QComboBox()
        value_combobox.setEnabled(False)  # Блокируем до выбора столбца
        filter_row.layout().addWidget(value_combobox)

        # Кнопка удаления фильтра
        delete_button = QPushButton("Удалить")
        delete_button.clicked.connect(lambda: self.remove_filter(filter_row))
        filter_row.layout().addWidget(delete_button)

        self.filter_frame.layout().addWidget(filter_row)

        # Обработчики изменений
        def on_column_changed():
            column = column_combobox.currentText()
            value_combobox.setEnabled(bool(column))  # Активируем только при выбранном столбце
            self.update_filter_values(column_combobox, value_combobox)

        column_combobox.currentTextChanged.connect(on_column_changed)

    def remove_filter(self, filter_row):
        """
        Удаляет фильтр из интерфейса.
        """
        filter_row.setParent(None)

    def update_filter_values(self, column_combobox, value_combobox):
        """
        Обновляет значения в выпадающем списке для фильтра.
        """
        column = column_combobox.currentText()
        if column:
            unique_values = self.current_df[column].dropna().unique()
            value_combobox.clear()
            value_combobox.addItems(map(str, unique_values))

    def apply_filters(self):
        """
        Применяет все активные фильтры к данным.
        """
        if self.current_df.empty:
            QMessageBox.warning(self, "Предупреждение", "Данные отсутствуют для фильтрации.")
            return

        filtered_df = self.current_df.copy()
        for i in range(self.filter_frame.layout().count()):
            filter_row = self.filter_frame.layout().itemAt(i).widget()
            if isinstance(filter_row, QFrame):
                column_combobox = filter_row.layout().itemAt(1).widget()
                value_combobox = filter_row.layout().itemAt(3).widget()

                column = column_combobox.currentText()
                value = value_combobox.currentText()

                # Пропускаем незаполненные фильтры
                if not column or not value:
                    continue

                if column not in self.current_df.columns:
                    QMessageBox.warning(self, "Ошибка", f"Столбец '{column}' отсутствует в данных.")
                    continue

                filtered_df = filtered_df[filtered_df[column].astype(str) == value]

        self.filtered_df = filtered_df
        self.update_table(filtered_df)
        QMessageBox.information(self, "Успех", f"Фильтры применены. Найдено {len(filtered_df)} строк.")

    def save_filtered_to_excel(self):
        """
        Сохраняет отфильтрованные данные в Excel.
        """
        if self.filtered_df is None or self.filtered_df.empty:
            QMessageBox.critical(self, "Ошибка", "Нет данных для сохранения.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как Excel", "", "Excel Files (*.xlsx)"
        )
        if not save_path:
            return

        try:
            self.filtered_df.to_excel(save_path, index=False, engine="openpyxl")
            QMessageBox.information(self, "Успех", f"Данные успешно сохранены в файл:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")

    def save_filtered_to_csv(self):
        """
        Сохраняет отфильтрованные данные в CSV.
        """
        if self.filtered_df is None or self.filtered_df.empty:
            QMessageBox.critical(self, "Ошибка", "Нет данных для сохранения.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как CSV", "", "CSV Files (*.csv)"
        )
        if not save_path:
            return

        try:
            self.filtered_df.to_csv(save_path, index=False, encoding="utf-8-sig")
            QMessageBox.information(self, "Успех", f"Данные успешно сохранены в файл:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")

    def export_csv_to_excel(self):
        """
        Конвертирует CSV в Excel.
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите CSV файл", "", "CSV Files (*.csv);;Все файлы (*.*)"
        )
        if not file_path:
            QMessageBox.critical(self, "Ошибка", "Файл не выбран.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как Excel", "", "Excel Files (*.xlsx)"
        )
        if not save_path:
            return

        try:
            df = pd.read_csv(file_path)
            df.to_excel(save_path, index=False, engine="openpyxl")
            QMessageBox.information(self, "Успех", "Файл успешно сохранён как Excel!")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось конвертировать файл: {e}")

    def export_file(self):
        """
        Экспорт данных в выбранный формат.
        """
        if not hasattr(self, "current_df") or self.current_df is None:
            QMessageBox.critical(self, "Ошибка", "Сначала загрузите и просмотрите данные.")
            return

        selected_format = self.format_combo.currentText()
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как",
            "", f"{selected_format.upper()} Files (*.{selected_format.lower()})"
        )
        if save_path:
            try:
                export_to_format(self.current_df, selected_format, save_path)
                QMessageBox.information(self, "Успех", f"Файл успешно сохранён в формате {selected_format}!")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelToCsvApp()
    window.show()
    sys.exit(app.exec_())