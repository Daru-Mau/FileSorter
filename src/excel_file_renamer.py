#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel-Based File Renamer

A GUI application for renaming files based on matching data in Excel spreadsheets.
"""

import os
import shutil
import re
import pandas as pd
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar, BOTH, X, Y, LEFT, RIGHT, END, W, SUNKEN, VERTICAL


class ExcelFileRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel-Based File Renamer")
        self.root.geometry("1100x700")
        self.root.minsize(900, 600)

        self.source_folder = StringVar()
        self.excel_file_path = StringVar()
        self.search_term = StringVar()

        self.selected_file = None
        self.file_list_data = []
        self.excel_data = None
        self.name_column = StringVar()
        self.id_column = StringVar()
        self.date_column = StringVar()

        self.setup_ui()

    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # Top section - file and excel selection
        top_frame = ttk.LabelFrame(
            main_frame, text="File Selection", padding="10")
        top_frame.pack(fill=X, pady=5)

        # Source folder
        ttk.Label(top_frame, text="Source Folder:").grid(
            row=0, column=0, sticky=W, pady=5)
        ttk.Entry(top_frame, textvariable=self.source_folder,
                  width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="Browse...", command=self.browse_source).grid(
            row=0, column=2, padx=5, pady=5)

        # Excel file
        ttk.Label(top_frame, text="Excel File:").grid(
            row=1, column=0, sticky=W, pady=5)
        ttk.Entry(top_frame, textvariable=self.excel_file_path,
                  width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="Browse...", command=self.browse_excel).grid(
            row=1, column=2, padx=5, pady=5)
        ttk.Button(top_frame, text="Load Excel", command=self.load_excel_data).grid(
            row=1, column=3, padx=5, pady=5)

        # Excel column mapping frame
        excel_map_frame = ttk.LabelFrame(
            main_frame, text="Excel Column Mapping", padding="10")
        excel_map_frame.pack(fill=X, pady=5)

        # Name column mapping
        ttk.Label(excel_map_frame, text="Name Column:").grid(
            row=0, column=0, sticky=W, pady=5)
        self.name_column_combo = ttk.Combobox(
            excel_map_frame, textvariable=self.name_column, width=20)
        self.name_column_combo.grid(row=0, column=1, padx=5, pady=5)

        # ID column mapping
        ttk.Label(excel_map_frame, text="ID Column:").grid(
            row=0, column=2, sticky=W, pady=5)
        self.id_column_combo = ttk.Combobox(
            excel_map_frame, textvariable=self.id_column, width=20)
        self.id_column_combo.grid(row=0, column=3, padx=5, pady=5)

        # Date column mapping
        ttk.Label(excel_map_frame, text="Date Column:").grid(
            row=0, column=4, sticky=W, pady=5)
        self.date_column_combo = ttk.Combobox(
            excel_map_frame, textvariable=self.date_column, width=20)
        self.date_column_combo.grid(row=0, column=5, padx=5, pady=5)

        # Middle section - paned window for files and excel data
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill=BOTH, expand=True, pady=5)

        # Left panel - Files list
        files_frame = ttk.LabelFrame(paned_window, text="Files", padding="10")
        paned_window.add(files_frame, weight=1)

        # Search box for files
        search_frame = ttk.Frame(files_frame)
        search_frame.pack(fill=X, pady=5)

        ttk.Label(search_frame, text="Search:").pack(side=LEFT, padx=5)
        ttk.Entry(search_frame, textvariable=self.search_term,
                  width=30).pack(side=LEFT, padx=5)
        ttk.Button(search_frame, text="Search",
                   command=self.search_files).pack(side=LEFT, padx=5)
        ttk.Button(search_frame, text="Scan Files",
                   command=self.scan_files).pack(side=LEFT, padx=5)

        # Files listbox with scrollbar
        self.files_listbox = tk.Listbox(files_frame, width=40, height=20)
        self.files_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        files_scrollbar = ttk.Scrollbar(
            files_frame, orient=VERTICAL, command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)
        files_scrollbar.pack(side=RIGHT, fill=Y)

        # Bind listbox selection event
        self.files_listbox.bind('<<ListboxSelect>>', self.on_file_select)

        # Right panel - Excel data display
        excel_frame = ttk.LabelFrame(
            paned_window, text="Excel Data", padding="10")
        paned_window.add(excel_frame, weight=2)

        # Excel data treeview
        self.excel_tree = ttk.Treeview(excel_frame, show="headings")

        # Scrollbars for excel tree
        excel_y_scrollbar = ttk.Scrollbar(
            excel_frame, orient=VERTICAL, command=self.excel_tree.yview)
        excel_x_scrollbar = ttk.Scrollbar(
            excel_frame, orient=tk.HORIZONTAL, command=self.excel_tree.xview)
        self.excel_tree.configure(
            yscrollcommand=excel_y_scrollbar.set, xscrollcommand=excel_x_scrollbar.set)

        # Pack excel tree and scrollbars
        self.excel_tree.pack(side=LEFT, fill=BOTH, expand=True)
        excel_y_scrollbar.pack(side=RIGHT, fill=Y)
        excel_x_scrollbar.pack(side=tk.BOTTOM, fill=X)

        # Manual rename section
        manual_frame = ttk.LabelFrame(
            main_frame, text="Manual Rename", padding="10")
        manual_frame.pack(fill=X, pady=5)

        # Manual rename entry
        ttk.Label(manual_frame, text="Custom Filename:").grid(
            row=0, column=0, sticky=W, pady=5)
        self.manual_filename = StringVar()
        self.manual_entry = ttk.Entry(
            manual_frame, textvariable=self.manual_filename, width=40)
        self.manual_entry.grid(row=0, column=1, padx=5, pady=5, sticky=W)

        # Keep extension checkbox
        self.keep_extension = tk.BooleanVar(value=True)
        ttk.Checkbutton(manual_frame, text="Keep original extension", variable=self.keep_extension).grid(
            row=0, column=2, padx=5, pady=5, sticky=W)

        # Manual rename button
        self.manual_rename_button = ttk.Button(manual_frame, text="Apply Manual Rename",
                                               command=self.manual_rename, state=tk.DISABLED)
        self.manual_rename_button.grid(row=0, column=3, padx=5, pady=5)

        # Bottom section - actions
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=X, pady=10)

        # Details label
        self.details_label = ttk.Label(
            action_frame, text="", wraplength=800, justify=tk.LEFT)
        self.details_label.pack(side=LEFT, fill=X, expand=True, padx=5)

        # Rename button
        self.rename_button = ttk.Button(
            action_frame, text="Rename Selected File", command=self.rename_file, state=tk.DISABLED)
        self.rename_button.pack(side=RIGHT, padx=5)

        # Status bar
        self.status_var = StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(
            main_frame, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        status_bar.pack(side=tk.BOTTOM, fill=X)

    def browse_source(self):
        """Browse for source folder"""
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder.set(folder)
            self.scan_files()

    def browse_excel(self):
        """Browse for Excel file"""
        file_path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx;*.xls"),
            ("All files", "*.*")
        ])
        if file_path:
            self.excel_file_path.set(file_path)

    def load_excel_data(self):
        """Load data from Excel file"""
        excel_path = self.excel_file_path.get()
        if not excel_path:
            messagebox.showerror("Error", "Please select an Excel file")
            return

        try:
            # Load Excel data
            self.excel_data = pd.read_excel(excel_path)

            # Clear current treeview
            for column in self.excel_tree["columns"]:
                self.excel_tree.heading(column, text="")

            for item in self.excel_tree.get_children():
                self.excel_tree.delete(item)

            # Configure columns
            columns = list(self.excel_data.columns)
            self.excel_tree["columns"] = columns

            # Update column mapping dropdown options
            self.name_column_combo["values"] = columns
            self.id_column_combo["values"] = columns
            self.date_column_combo["values"] = columns

            # Set default values if possible
            name_matches = [col for col in columns if 'name' in col.lower()]
            id_matches = [col for col in columns if 'id' in col.lower()]
            date_matches = [col for col in columns if 'date' in col.lower()]

            if name_matches:
                self.name_column.set(name_matches[0])
            elif columns:
                self.name_column.set(columns[0])

            if id_matches:
                self.id_column.set(id_matches[0])
            elif len(columns) > 1:
                self.id_column.set(columns[1])

            if date_matches:
                self.date_column.set(date_matches[0])

            # Set up headings
            for col in columns:
                self.excel_tree.heading(col, text=col)
                # Adjust column width based on content
                max_width = max([len(str(self.excel_data[col].iloc[i])) for i in range(
                    min(10, len(self.excel_data)))] + [len(col)])
                self.excel_tree.column(col, width=max_width * 10)

            # Insert data rows
            for i, row in self.excel_data.iterrows():
                values = [row[col] for col in columns]
                self.excel_tree.insert("", END, values=values)

            self.status_var.set(
                f"Loaded {len(self.excel_data)} rows from Excel")

        except Exception as e:
            messagebox.showerror(
                "Error", f"Failed to load Excel file: {str(e)}")
            self.status_var.set("Error loading Excel data")

    def scan_files(self):
        """Scan files in the source folder"""
        source = self.source_folder.get()
        if not source:
            messagebox.showerror("Error", "Please select a source folder")
            return

        # Clear current data
        self.file_list_data = []
        self.files_listbox.delete(0, END)

        try:
            # Get all files in the folder
            files = list(Path(source).glob("*"))

            # Filter out directories
            files = [f for f in files if f.is_file()]

            # Sort files by name
            files.sort(key=lambda x: x.name.lower())

            # Store file data and update listbox
            for file_path in files:
                file_stat = file_path.stat()
                mod_time = datetime.fromtimestamp(
                    file_stat.st_mtime).strftime('%Y-%m-%d %H:%M')

                # Store file data
                file_data = {
                    "path": file_path,
                    "name": file_path.name,
                    "size": file_stat.st_size,
                    "mod_time": file_stat.st_mtime,
                    "mod_time_formatted": mod_time,
                    "extension": file_path.suffix.lower()
                }
                self.file_list_data.append(file_data)

                # Add to listbox
                self.files_listbox.insert(END, file_path.name)

            self.status_var.set(f"Found {len(files)} files")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set("Error scanning files")

    def search_files(self):
        """Search files by name"""
        search_text = self.search_term.get().lower()
        if not search_text:
            # If search is empty, show all files
            self.scan_files()
            return

        # Clear listbox
        self.files_listbox.delete(0, END)

        # Filter and display matching files
        for file_data in self.file_list_data:
            if search_text in file_data["name"].lower():
                self.files_listbox.insert(END, file_data["name"])

        self.status_var.set(f"Found {self.files_listbox.size()} matches")

    def on_file_select(self, event):
        """Handle file selection from listbox"""
        if not self.files_listbox.curselection():
            return

        # Get selected file data
        selected_index = self.files_listbox.curselection()[0]
        selected_filename = self.files_listbox.get(selected_index)

        # Find matching file data
        for file_data in self.file_list_data:
            if file_data["name"] == selected_filename:
                self.selected_file = file_data
                break

        # Enable manual rename button when a file is selected
        self.manual_rename_button.config(state=tk.NORMAL)

        # Populate manual rename field with current filename without extension
        filename_without_ext = os.path.splitext(selected_filename)[0]
        self.manual_filename.set(filename_without_ext)

        # Find matching entry in Excel
        if self.excel_data is not None and self.name_column.get():
            # Get filename without extension for matching
            filename_without_ext = os.path.splitext(selected_filename)[0]

            # Try different matching strategies
            match_found = False

            # Clear previous selection
            for item in self.excel_tree.selection():
                self.excel_tree.selection_remove(item)

            # Strategy 1: Exact match
            for i, item in enumerate(self.excel_tree.get_children()):
                values = self.excel_tree.item(item)["values"]
                col_index = list(self.excel_data.columns).index(
                    self.name_column.get())
                excel_value = str(values[col_index])

                if excel_value == filename_without_ext or excel_value == selected_filename:
                    self.excel_tree.selection_set(item)
                    self.excel_tree.see(item)
                    match_found = True
                    self.update_details(item)
                    break

            # Strategy 2: Contains match (if no exact match found)
            if not match_found:
                for i, item in enumerate(self.excel_tree.get_children()):
                    values = self.excel_tree.item(item)["values"]
                    col_index = list(self.excel_data.columns).index(
                        self.name_column.get())
                    excel_value = str(values[col_index])

                    if filename_without_ext in excel_value or excel_value in filename_without_ext:
                        self.excel_tree.selection_set(item)
                        self.excel_tree.see(item)
                        match_found = True
                        self.update_details(item)
                        break

            if match_found:
                self.status_var.set(f"Found match for {selected_filename}")
                self.rename_button.config(state=tk.NORMAL)
            else:
                self.status_var.set(f"No match found for {selected_filename}")
                self.details_label.config(
                    text=f"No match found in Excel data for {selected_filename}")
                self.rename_button.config(state=tk.DISABLED)
        else:
            self.status_var.set(
                "Excel data not loaded or columns not configured")

    def update_details(self, item):
        """Update details label with information about the selected file and Excel match"""
        if not self.selected_file:
            return

        # Get Excel row data
        values = self.excel_tree.item(item)["values"]
        columns = list(self.excel_data.columns)  # type: ignore
        excel_data_dict = {columns[i]: values[i] for i in range(len(columns))}

        # Format file details
        file_details = (
            f"File: {self.selected_file['name']}\n"
            f"Size: {self.format_size(self.selected_file['size'])}\n"
            f"Modified: {self.selected_file['mod_time_formatted']}\n\n"
        )

        # Format Excel match details
        excel_details = "Excel Match:\n"
        for col, val in excel_data_dict.items():
            excel_details += f"{col}: {val}\n"

        # Show proposed new filename
        id_val = excel_data_dict.get(self.id_column.get(), "")
        if id_val:
            extension = self.selected_file["extension"]
            proposed_name = f"{id_val}{extension}"
            excel_details += f"\nProposed new filename: {proposed_name}"

        # Update details label
        self.details_label.config(text=file_details + excel_details)

    def format_size(self, size_bytes):
        """Format file size in human-readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"

    def rename_file(self):
        """Rename the selected file based on Excel data"""
        if not self.selected_file or not self.excel_tree.selection():
            messagebox.showerror("Error", "No file or Excel entry selected")
            return

        # Get Excel row data
        selected_item = self.excel_tree.selection()[0]
        values = self.excel_tree.item(selected_item)["values"]
        columns = list(self.excel_data.columns)  # type: ignore
        excel_data_dict = {columns[i]: values[i] for i in range(len(columns))}

        # Get ID value for new name
        id_val = excel_data_dict.get(self.id_column.get(), "")
        if not id_val:
            messagebox.showerror(
                "Error", f"No ID value found in the '{self.id_column.get()}' column")
            return

        # Format ID value to ensure it's a valid filename
        id_val = str(id_val).strip()
        # Replace invalid filename chars
        id_val = re.sub(r'[\\/*?:"<>|]', '_', id_val)

        # Create new filename
        extension = self.selected_file["extension"]
        new_filename = f"{id_val}{extension}"

        # Check if source and new names are the same
        if new_filename == self.selected_file["name"]:
            messagebox.showinfo(
                "Info", "The file already has the correct name")
            return

        # Create full paths
        source_path = self.selected_file["path"]
        dest_folder = self.source_folder.get()  # Same folder for now
        dest_path = Path(dest_folder) / new_filename

        # Check if destination file already exists
        if dest_path.exists():
            confirm = messagebox.askyesno(
                "Confirm Overwrite",
                f"File {new_filename} already exists. Overwrite?")
            if not confirm:
                return

        try:
            # Rename file
            os.rename(source_path, dest_path)

            # Update file data
            self.selected_file["name"] = new_filename
            self.selected_file["path"] = dest_path

            # Update listbox
            selected_index = self.files_listbox.curselection()[0]
            self.files_listbox.delete(selected_index)
            self.files_listbox.insert(selected_index, new_filename)
            self.files_listbox.selection_set(selected_index)

            # Update status
            self.status_var.set(f"Renamed file to {new_filename}")
            messagebox.showinfo("Success", f"Renamed file to {new_filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file: {str(e)}")
            self.status_var.set("Error renaming file")

    def manual_rename(self):
        """Rename the selected file using a custom name"""
        if not self.selected_file:
            messagebox.showerror("Error", "No file selected")
            return

        # Get the custom filename
        custom_name = self.manual_filename.get().strip()
        if not custom_name:
            messagebox.showerror("Error", "Please enter a custom filename")
            return

        # Validate filename (no invalid characters)
        if re.search(r'[\\/*?:"<>|]', custom_name):
            messagebox.showerror(
                "Error", "Filename contains invalid characters (\\/*?:\"<>|)")
            return

        # Add extension if needed
        extension = self.selected_file["extension"]
        if self.keep_extension.get() and not custom_name.lower().endswith(extension.lower()):
            new_filename = f"{custom_name}{extension}"
        else:
            new_filename = custom_name

        # Create full paths
        source_path = self.selected_file["path"]
        dest_folder = self.source_folder.get()
        dest_path = Path(dest_folder) / new_filename

        # Check if destination file already exists
        if dest_path.exists():
            confirm = messagebox.askyesno(
                "Confirm Overwrite",
                f"File {new_filename} already exists. Overwrite?")
            if not confirm:
                return

        try:
            # Rename file
            os.rename(source_path, dest_path)

            # Update file data
            self.selected_file["name"] = new_filename
            self.selected_file["path"] = dest_path

            # Update listbox
            selected_index = self.files_listbox.curselection()[0]
            self.files_listbox.delete(selected_index)
            self.files_listbox.insert(selected_index, new_filename)
            self.files_listbox.selection_set(selected_index)

            # Clear the manual filename entry
            self.manual_filename.set("")

            # Update status
            self.status_var.set(f"Manually renamed file to {new_filename}")
            messagebox.showinfo("Success", f"Renamed file to {new_filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file: {str(e)}")
            self.status_var.set("Error renaming file")


def main():
    root = tk.Tk()
    app = ExcelFileRenamerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
