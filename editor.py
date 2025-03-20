from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox
from PyQt5.QtCore import Qt


class CellEditor(QDialog):
    def __init__(self, parent, tree, current_df):
        super().__init__(parent)
        self.tree = tree
        self.current_df = current_df
        self.selected_item = None
        self.col_index = None

        self.setWindowTitle("Редактирование ячейки")
        self.setModal(True)
        self.setup_ui()

    def setup_ui(self):
        """
        Настраивает интерфейс окна редактирования.
        """
        layout = QVBoxLayout(self)

        self.label = QLabel("Новое значение:")
        layout.addWidget(self.label)

        self.new_value_entry = QLineEdit()
        layout.addWidget(self.new_value_entry)

        self.save_button = QPushButton("Сохранить")
        self.save_button.clicked.connect(self.save_new_value)
        layout.addWidget(self.save_button)

    def edit_cell(self, index):
        """
        Открывает окно редактирования для выбранной ячейки.
        """
        self.selected_item = index
        self.col_index = index.column()

        current_value = self.tree.model().data(index, Qt.DisplayRole)
        self.new_value_entry.setText(current_value)

        self.exec_()

    def save_new_value(self):
        """
        Сохраняет новое значение в ячейке.
        """
        new_value = self.new_value_entry.text()

        if not new_value:
            QMessageBox.warning(self, "Ошибка", "Значение не может быть пустым.")
            return

        model = self.tree.model()
        model.setData(self.selected_item, new_value, Qt.DisplayRole)

        row = self.selected_item.row()
        self.current_df.iat[row, self.col_index] = new_value

        self.accept()