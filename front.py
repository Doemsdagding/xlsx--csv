import tkinter as tk
from tkinter import ttk
import json
import os


class EditRulesWindow:
    """
    A window for editing table rules and saving them to a file.
    Allows adding, deleting, and editing rows in a table, with
        persistent storage in the user's AppData folder.
    """
    def __init__(self, headers, save_file, title):
        """
        Initialize the EditRulesWindow.
        Args:
            headers (list): List of column headers for the table.
            save_file (str): Filename for saving the table data.
            title (str): Title of the window.
        """
        # Initialize the main window
        self.headers = headers
        self.save_file = self.get_user_data_path(save_file)
        self.rows = []
        self.table_window = tk.Toplevel()
        self.table_window.title(title)
        self.table_window.geometry("600x400")

        frame = tk.Frame(self.table_window)
        frame.pack(fill=tk.BOTH, expand=True)

        # Create a Treeview widget for the table
        self.tree = ttk.Treeview(frame, columns=self.headers, show="headings")
        for header in self.headers:
            self.tree.heading(header, text=header)
            self.tree.column(header, width=100, anchor="center")

        # Add a vertical scrollbar to the treeview
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Buttons for adding, deleting
        self.add_row_button = tk.Button(
            self.table_window, text="Add Row", command=self.add_row
        )
        self.add_row_button.pack(pady=10)

        self.delete_row_button = tk.Button(
            self.table_window, text="Delete Row", command=self.delete_row
        )
        self.delete_row_button.pack(pady=10)

        # Variables for tracking the editing state
        self.editing_entry = None
        self.editing_item = None
        self.editing_column_index = None

        # Bind events for editing cells
        self.tree.bind("<Button-1>", self.edit_cell)
        frame.bind("<Button-1>", self.save_edit_on_click)

        # Button to save all rows to the file
        if self.save_file and os.path.exists(self.save_file):
            self.load_rows()
        save_button = tk.Button(
            self.table_window, text="Save Rows", command=self.save_rows
        )
        save_button.pack(pady=10)

    def get_user_data_path(self, filename):
        """
        Get the full path for saving user data in the AppData folder.
        Args:
            filename (str): The filename to use for saving data.
        Returns:
            str: The full path to the file in AppData.
        """
        appdata = os.getenv("APPDATA") or os.path.expanduser("~")
        folder = os.path.join(appdata, "xlsx_csv_repo")
        # Create the xlsx_csv_repo folder if it doesn't exist
        os.makedirs(folder, exist_ok=True)
        return os.path.join(folder, filename)

    def add_row(self):
        """
        Add a new empty row to the table.
        """
        self.tree.insert("", "end", values=["" for _ in self.headers])

    def delete_row(self):
        """
        Delete the selected rows from the table.
        """
        selected_items = self.tree.selection()
        for item in selected_items:
            self.tree.delete(item)

    def save_edit(self, event=None):
        """
        Save the current edit in the table cell, if any.
        """
        if (
            self.editing_entry
            and self.editing_item is not None
            and self.editing_column_index is not None
        ):
            # Retrieve value from entry widget and update the treeview
            new_value = self.editing_entry.get()
            values = list(self.tree.item(self.editing_item, "values"))
            values[self.editing_column_index] = new_value
            self.tree.item(self.editing_item, values=values)
            self.editing_entry.destroy()
            self.editing_entry = None
            self.editing_item = None
            self.editing_column_index = None

    def edit_cell(self, event):
        """
        Start editing a cell in the table when clicked.
        """
        # Save any currently edited cell before starting a new edit
        if self.editing_entry:
            self.save_edit()

        # Get the clicked item and column
        selected_item = self.tree.selection()

        # if there is a selected item, start editing
        if selected_item:
            item = selected_item[0]
            column = self.tree.identify_column(event.x)
            col_index = int(column.replace("#", "")) - 1
            # Check if the column index is valid
            if col_index >= 0:
                # Create the entry widget for editing
                x, y, width, height = self.tree.bbox(item, column)
                self.editing_entry = tk.Entry(self.tree)
                self.editing_entry.place(x=x, y=y, width=width, height=height)
                self.editing_entry.insert(
                    0, self.tree.item(item, "values")[col_index]
                )
                # Bind the location of the field to use while saving
                self.editing_item = item
                self.editing_column_index = col_index

                # Bind the entry to save on return key
                self.editing_entry.bind("<Return>", self.save_edit)
                self.editing_entry.focus()

    def save_edit_on_click(self, event):
        """
        Save the current edit when clicking outside the editing cell.
        """
        if self.editing_entry:
            self.save_edit()

    def save_rows(self):
        """
        Save all rows to the file, including any currently edited cell.
        """
        # Save any currently edited field before saving all rows
        if self.editing_entry:
            self.save_edit()
        # Save all rows to file
        if self.save_file:
            rows = [self.tree.item(
                item, "values") for item in self.tree.get_children()]
            with open(self.save_file, "w", encoding="utf-8") as f:
                json.dump(rows, f)

    def load_rows(self):
        """
        Load rows from the file and insert them into the table.
        """
        # Load rows from file
        with open(self.save_file, "r", encoding="utf-8") as f:
            rows = json.load(f)
        for row in rows:
            self.tree.insert("", "end", values=row)

class EditDefaultWindow(EditRulesWindow):
    """
    A specialized EditRulesWindow that hides add row and delete row buttons.
    Used for editing default rule values.
    """
    def __init__(self, headers, save_file, title):
        """
        Initialize the EditDefaultWindow.
        Args:
            headers (list): List of column headers for the table.
            save_file (str): Filename for saving the table data.
            title (str): Title of the window.
        """
        super().__init__(headers, save_file, title)
        # Hide add row and delete row buttons
        self.add_row_button.pack_forget()
        self.delete_row_button.pack_forget()


class MainWindow:
    """
    The main window for the XLSX to CSV Converter.
    """
    def __init__(self):
        """
        Initialize the main window and its buttons to the rule windows.
        """
        # Set window info
        self.window = tk.Tk()
        self.window.title("XLSX_to_CSV_Converter")
        self.window.geometry("400x300")

        label = tk.Label(
            self.window, text="XLSX to CSV Converter", font=("Arial", 16)
        )
        label.pack(pady=20)

        # Create buttons to open EditRulesWindow and EditDefaultWindow
        headers = ["Header1", "Header2", "Header3"]
        button_rules = tk.Button(
            self.window,
            text="Edit rules",
            command=lambda: EditRulesWindow(
                headers, save_file="rules.json", title="Edit Rules"
            ),
        )
        button_rules.pack(pady=10)

        button_default = tk.Button(
            self.window,
            text="Edit default",
            command=lambda: EditDefaultWindow(
                headers, save_file="default.json", title="Edit Default"
            ),
        )
        button_default.pack(pady=10)

    def run(self):
        """
        Start the Tkinter main loop for the application.
        """
        self.window.mainloop()

if __name__ == "__main__":
    app = MainWindow()
    app.run()