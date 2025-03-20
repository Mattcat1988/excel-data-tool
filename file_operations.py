import sqlite3
import os
import yaml
from tabulate import tabulate
import pandas as pd
import re
from PyQt5.QtWidgets import QFileDialog, QMessageBox


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
    def evaluate_cell(value):
        try:
            if isinstance(value, str):
                if value.startswith('='):
                    formula = value[1:]
                    formula = re.sub(r'SUM\((.*?)\)', r'sum([\1])', formula, flags=re.IGNORECASE)
                    return eval(formula)
                return value
            elif isinstance(value, (int, float)):
                return value
            return value
        except Exception:
            return value

    return df.map(evaluate_cell)


def eval_formula(formula):
    try:
        if formula.startswith('='):
            if formula.upper().startswith('=SUM'):
                numbers = re.findall(r'\d+', formula)
                return sum(map(int, numbers))
            return eval(formula[1:])
    except Exception as e:
        return f"Ошибка в формуле: {e}"


def convert_csv_to_excel(csv_path, excel_path):
    try:
        df = pd.read_csv(csv_path)
        df.to_excel(excel_path, index=False, encoding='utf-8-sig')
        QMessageBox.information(None, "Успех", f"Файл успешно сохранён как Excel:\n{excel_path}")
    except Exception as e:
        QMessageBox.critical(None, "Ошибка", f"Не удалось конвертировать файл: {e}")


def export_csv_to_excel(parent):
    file_path, _ = QFileDialog.getOpenFileName(
        parent, "Выберите CSV файл", "", "CSV Files (*.csv);;Все файлы (*.*)"
    )
    if not file_path:
        QMessageBox.critical(parent, "Ошибка", "Файл не выбран.")
        return

    save_path, _ = QFileDialog.getSaveFileName(
        parent, "Сохранить как Excel", "", "Excel Files (*.xlsx)"
    )
    if not save_path:
        return

    try:
        df = pd.read_csv(file_path)
        df.to_excel(save_path, index=False, engine="openpyxl")
        QMessageBox.information(parent, "Успех", "Файл успешно сохранён как Excel!")
    except Exception as e:
        QMessageBox.critical(parent, "Ошибка", f"Не удалось конвертировать файл: {e}")