import tkinter as tk


class CellEditor:
    def __init__(self, parent, tree, current_df):
        self.parent = parent
        self.tree = tree
        self.current_df = current_df

        # Пример добавления адаптивного размещения
        self.parent.columnconfigure(0, weight=1)
        self.parent.rowconfigure(0, weight=1)
        self.tree.grid(sticky="nsew")
    def edit_cell(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return

        selected_item = selected_item[0]
        col = self.tree.identify_column(event.x)[1:]  # Определяем колонку
        col_index = int(col) - 1  # Преобразуем номер колонки в индекс
        current_value = self.tree.item(selected_item)["values"][col_index]

        # Открываем окно редактирования
        self.open_edit_window(selected_item, col_index, current_value)



    def open_edit_window(self, selected_item, col_index, current_value):
        edit_window = tk.Toplevel(self.parent)
        edit_window.title("Редактирование ячейки")

        tk.Label(edit_window, text="Новое значение:").pack(pady=5)

        new_value_entry = tk.Entry(edit_window)
        new_value_entry.pack(pady=5)
        new_value_entry.insert(0, current_value)

        def save_new_value():
            new_value = new_value_entry.get()

            # Обновляем данные в Treeview
            self.tree.item(selected_item, values=[
                new_value if i == col_index else val
                for i, val in enumerate(self.tree.item(selected_item)["values"])
            ])

            # Обновляем данные в DataFrame
            self.current_df.iat[int(selected_item[1:], 16) - 1, col_index] = new_value

            edit_window.destroy()

        tk.Button(edit_window, text="Сохранить", command=save_new_value).pack(pady=10)