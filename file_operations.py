import sqlite3
import os
import yaml
from tabulate import tabulate
import pandas as pd
import re
from tkinter import filedialog, messagebox


def process_data(df):
    return df.map(lambda x: round(x, 2) if isinstance(x, float) else x)


def export_to_format(df, format, save_path):
    if format == "CSV":
        df.to_csv(save_path, index=False, encoding="utf-8-sig")
    elif format == "JSON":
        df.to_json(save_path, orient="records", force_ascii=False, indent=4)
    elif format == "XML":
        export_to_xml(df, save_path)
    elif format == "Markdown":
        export_to_markdown(df, save_path)
    elif format == "YAML":
        export_to_yaml(df, save_path)
    elif format == "SQL":
        export_to_sql(df, save_path)
    elif format == "HTML":
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(df.to_html(index=False))


def export_to_xml(df, save_path):
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write('<root>\n')
        for _, row in df.iterrows():
            f.write('  <row>\n')
            for field, value in row.items():
                f.write(f'    <{field}>{value}</{field}>\n')
            f.write('  </row>\n')
        f.write('</root>')


def export_to_markdown(df, save_path):
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write(tabulate(df, headers='keys', tablefmt='pipe'))


def export_to_yaml(df, save_path):
    with open(save_path, 'w', encoding='utf-8') as f:
        yaml.dump(df.to_dict(orient="records"), f, allow_unicode=True, default_flow_style=False)


def export_to_sql(df, save_path):
    conn = sqlite3.connect(save_path)
    table_name = os.path.splitext(os.path.basename(save_path))[0]
    df.to_sql(table_name, conn, if_exists="replace", index=False)
    conn.close()



def evaluate_formulas(df):
    """
    Обрабатывает значения в DataFrame, вычисляя формулы или оставляя их неизменными.
    """
    def evaluate_cell(value):
        try:
            # Проверяем, является ли значение строкой
            if isinstance(value, str):
                # Если начинается с '=', предполагаем формулу
                if value.startswith('='):
                    formula = value[1:]
                    # Замена синтаксиса Excel на Python
                    formula = re.sub(r'SUM\((.*?)\)', r'sum([\1])', formula, flags=re.IGNORECASE)
                    return eval(formula)  # Осторожно с eval!
                # Иначе просто возвращаем текст
                return value
            elif isinstance(value, (int, float)):
                return value  # Числовые значения остаются неизменными
            return value
        except Exception:
            return value  # При ошибке возвращаем оригинальное значение

    return df.map(evaluate_cell)


def eval_formula(formula):
    """
    Эмулирует выполнение простых Excel-формул (например, "=SUM(1,2)").
    """
    try:
        if formula.startswith('='):
            # Пример простого парсинга формулы Excel
            if formula.upper().startswith('=SUM'):
                numbers = re.findall(r'\d+', formula)
                return sum(map(int, numbers))
            # Добавьте дополнительные правила для других функций Excel
            return eval(formula[1:])  # Осторожно с eval!
    except Exception as e:
        return f"Ошибка в формуле: {e}"

def convert_csv_to_excel(csv_path, excel_path):
    """
    Конвертирует CSV файл в Excel.
    """
    try:
        df = pd.read_csv(csv_path)
        df.to_excel(excel_path, index=False, encoding='utf-8-sig')
        messagebox.showinfo("Успех", f"Файл успешно сохранён как Excel:\n{excel_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось конвертировать файл: {e}")


# Включаем в интерфейс:
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
