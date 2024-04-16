import tkinter as tk
from tkinter import ttk
import tkinter.font as tkFont


class ExcelLikeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel-like Interface")

        self.tree = None
        self._setup_widgets()
        self._build_tree()

    def _setup_widgets(self):
        # Adding a frame for Treeview and scrollbar
        container = ttk.Frame(self.root)
        container.pack(fill='both', expand=True)

        # Creating Treeview widget
        self.tree = ttk.Treeview(columns=self._get_columns(), show="headings")
        self.tree.pack(side='left', fill='both', expand=True)

        # Vertical scrollbar
        vsb = ttk.Scrollbar(orient="vertical", command=self.tree.yview)
        vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=vsb.set)

        # Horizontal scrollbar
        hsb = ttk.Scrollbar(orient="horizontal", command=self.tree.xview)
        hsb.pack(side='bottom', fill='x')
        self.tree.configure(xscrollcommand=hsb.set)

        # Add and Delete buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill='x')
        add_button = ttk.Button(btn_frame, text="Add Row", command=self.add_row)
        add_button.pack(side='left')
        delete_button = ttk.Button(btn_frame, text="Delete Row", command=self.delete_row)
        delete_button.pack(side='left')

    def _build_tree(self):
        font = tkFont.Font(family="Helvetica", size=12)  # Crear una instancia de fuente
        for col in self._get_columns():
            self.tree.heading(col, text=col)
            # Usar la fuente para medir el ancho del texto y configurar el ancho de la columna
            self.tree.column(col, width=font.measure(col.title()) + 10)  # Agregar un pequeño buffer

        # Dummy data
        for i in range(10):
            self.tree.insert('', 'end', values=[f'Data {i + 1}' for _ in self._get_columns()])

        # Double-click to edit
        self.tree.bind("<Double-1>", self.on_double_click)

    def _get_columns(self):
        return ('Column 1', 'Column 2', 'Column 3', 'Column 4')

    def add_row(self):
        self.tree.insert('', 'end', values=[''] * len(self._get_columns()))

    def delete_row(self):
        selected_item = self.tree.selection()
        if selected_item:
            self.tree.delete(selected_item)

    def on_double_click(self, event):
        # Get selected item
        item = self.tree.identify('item', event.x, event.y)
        column = self.tree.identify_column(event.x)
        entry_edit = ttk.Entry(self.root)
        entry_edit.insert(0, self.tree.item(item, 'values')[int(column.replace('#', '')) - 1])
        entry_edit.place(x=event.x, y=event.y)

        def save_edit(event):
            self.tree.set(item, column=column, value=entry_edit.get())
            entry_edit.destroy()

        entry_edit.bind('<Return>', save_edit)
        entry_edit.bind('<FocusOut>', lambda e: entry_edit.destroy())
        entry_edit.focus_set()


def main():
    root = tk.Tk()
    app = ExcelLikeApp(root)
    root.mainloop()


main()
