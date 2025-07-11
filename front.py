import tkinter as tk
from tkinter import ttk

def edit_rules_window(headers):
    table_window = tk.Toplevel()
    table_window.title("Excel-like Table")
    table_window.geometry("600x400")

    # Create a frame for the table
    frame = tk.Frame(table_window)
    frame.pack(fill=tk.BOTH, expand=True)

    # Create the treeview (table)
    tree = ttk.Treeview(frame, columns=headers, show="headings")
    for header in headers:
        tree.heading(header, text=header)
        tree.column(header, width=100, anchor="center")
    tree.pack(fill=tk.BOTH, expand=True)

    # Add a button to add rows
    def add_row():
        tree.insert("", "end", values=["" for _ in headers])

    add_row_button = tk.Button(table_window, text="Add Row", command=add_row)
    add_row_button.pack(pady=10)

    # Add a button to delete selected rows
    def delete_row():
        selected_items = tree.selection()
        for item in selected_items:
            tree.delete(item)

    delete_row_button = tk.Button(table_window, text="Delete Row",
                                    command=delete_row)
    delete_row_button.pack(pady=10)

    # Enable cell editing trackers
    editing_entry = None
    editing_item = None
    editing_column_index = None

    def save_edit(event=None):
        nonlocal editing_entry, editing_item, editing_column_index
        if editing_entry and editing_item is not None \
        and editing_column_index is not None:
            new_value = editing_entry.get()
            values = list(tree.item(editing_item, "values"))
            values[editing_column_index] = new_value
            tree.item(editing_item, values=values)
            editing_entry.destroy()
            editing_entry = None
            editing_item = None
            editing_column_index = None

    def edit_cell(event):
        nonlocal editing_entry, editing_item, editing_column_index
        if editing_entry:
            save_edit()

        selected_item = tree.selection()
        if selected_item:
            item = selected_item[0]
            column = tree.identify_column(event.x)
            col_index = int(column.replace("#", "")) - 1
            if col_index >= 0:
                x, y, width, height = tree.bbox(item, column)
                editing_entry = tk.Entry(frame)
                editing_entry.place(x=x, y=y, width=width, height=height)
                editing_entry.insert(0, tree.item(item, "values")[col_index])
                editing_item = item
                editing_column_index = col_index

                editing_entry.bind("<Return>", save_edit)
                editing_entry.focus()

    def save_edit_on_click(event):
        nonlocal editing_entry
        if editing_entry:
            save_edit()

    tree.bind("<Button-1>", edit_cell)
    frame.bind("<Button-1>", save_edit_on_click)

def land_window():
    window = tk.Tk()
    window.title("XLSX_to_CSV_Converter")
    window.geometry("400x300")

    label = tk.Label(window, text="XLSX to CSV Converter",
        font=("Arial", 16))
    label.pack(pady=20)

    # Add a button to open the edit rules window
    headers = ["Header1", "Header2", "Header3"]
    button = tk.Button(window, text="Edit rules", 
                        command=lambda: edit_rules_window(headers))
    button.pack(pady=10)

    window.mainloop()

if __name__ == "__main__":
    land_window()