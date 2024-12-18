import tkinter as tk
from tkinter import font
from ui import ExcelToCsvApp  # Предполагается, что ваш основной класс интерфейса находится в модуле `ui`
import json
from pathlib import Path


def configure_scaling(root):
    """
    Настраивает масштабирование интерфейса в зависимости от операционной системы.
    """
    if root.tk.call("tk", "windowingsystem") == "win32":
        root.tk.call("tk", "scaling", 1.25)
    elif root.tk.call("tk", "windowingsystem") == "x11":
        root.tk.call("tk", "scaling", 1.0)
    elif root.tk.call("tk", "windowingsystem") == "aqua":
        root.tk.call("tk", "scaling", 2.0)


def set_default_fonts(root):
    """
    Устанавливает шрифты по умолчанию в зависимости от операционной системы.
    """
    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(size=12)

    if root.tk.call("tk", "windowingsystem") == "win32":
        default_font.configure(family="Segoe UI", size=10)
    elif root.tk.call("tk", "windowingsystem") == "x11":
        default_font.configure(family="Ubuntu", size=10)
    elif root.tk.call("tk", "windowingsystem") == "aqua":
        default_font.configure(family="San Francisco", size=12)


def save_window_config(root):
    """
    Сохраняет размеры и позицию окна в файл конфигурации.
    """
    config = {
        "width": root.winfo_width(),
        "height": root.winfo_height(),
        "x": root.winfo_x(),
        "y": root.winfo_y(),
    }
    config_path = Path.home() / ".app_config.json"
    with open(config_path, "w") as f:
        json.dump(config, f)


def load_window_config(root):
    """
    Загружает размеры и позицию окна из файла конфигурации.
    """
    config_path = Path.home() / ".app_config.json"
    if config_path.exists():
        with open(config_path, "r") as f:
            config = json.load(f)
        root.geometry(f"{config['width']}x{config['height']}+{config['x']}+{config['y']}")


if __name__ == "__main__":
    # Создание главного окна
    root = tk.Tk()
    root.title("Excel to CSV Converter")  # Примерное название, можно изменить
    

    # Загрузка конфигурации окна
    load_window_config(root)

    # Применение масштабирования и шрифтов
    configure_scaling(root)
    set_default_fonts(root)

    # Инициализация основного интерфейса
    app = ExcelToCsvApp(root)

    # Обработка события закрытия окна
    def on_close():
        save_window_config(root)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    

    # Запуск основного цикла
    
    root.mainloop()