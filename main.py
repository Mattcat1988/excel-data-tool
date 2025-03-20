import sys
import json
from pathlib import Path
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import QSettings, QRect, Qt
from PyQt5.QtGui import QFont
from ui import ExcelToCsvApp  # Основной класс интерфейса


def configure_scaling():
    """
    Настраивает масштабирование интерфейса в зависимости от операционной системы.
    """
    if sys.platform == "win32":
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)  # Включаем масштабирование для Windows
    elif sys.platform == "linux":
        pass  # Обычно не требуется дополнительных настроек для Linux
    elif sys.platform == "darwin":
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)  # Включаем масштабирование для macOS


def set_default_fonts(app):
    """
    Устанавливает шрифты по умолчанию в зависимости от операционной системы.
    """
    default_font = QFont()
    default_font.setPointSize(12)

    if sys.platform == "win32":
        default_font.setFamily("Segoe UI")
        default_font.setPointSize(10)
    elif sys.platform == "linux":
        default_font.setFamily("Ubuntu")
        default_font.setPointSize(10)
    elif sys.platform == "darwin":
        default_font.setFamily("San Francisco")
        default_font.setPointSize(12)

    app.setFont(default_font)


def save_window_config(window):
    """
    Сохраняет размеры и позицию окна в файл конфигурации.
    """
    config = {
        "width": window.width(),
        "height": window.height(),
        "x": window.x(),
        "y": window.y(),
    }
    config_path = Path.home() / ".app_config.json"
    with open(config_path, "w") as f:
        json.dump(config, f)


def load_window_config(window):
    """
    Загружает размеры и позицию окна из файла конфигурации.
    """
    config_path = Path.home() / ".app_config.json"
    if config_path.exists():
        with open(config_path, "r") as f:
            config = json.load(f)
        window.setGeometry(QRect(config['x'], config['y'], config['width'], config['height']))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to CSV Converter")  # Примерное название, можно изменить
        self.app = ExcelToCsvApp()  # Создаем экземпляр ExcelToCsvApp
        self.setCentralWidget(self.app)  # Устанавливаем его как центральный виджет

    def closeEvent(self, event):
        """
        Обработка события закрытия окна.
        """
        save_window_config(self)
        event.accept()


if __name__ == "__main__":
    # Настраиваем масштабирование до создания QApplication
    configure_scaling()

    # Создание приложения
    app = QApplication(sys.argv)

    # Применение шрифтов
    set_default_fonts(app)

    # Создание главного окна
    window = MainWindow()

    # Загрузка конфигурации окна
    load_window_config(window)

    # Показ окна
    window.show()

    # Запуск основного цикла
    sys.exit(app.exec_())