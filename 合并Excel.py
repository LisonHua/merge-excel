import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import xlrd.compdoc

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel合并工具V1.0 作者：LsionHua")
        self.root.geometry("1060x550")  # 设置窗口尺寸

        self.selected_files = []
        self.selected_sheets = []

        self.file_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=60, height=25)
        self.sheet_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=60, height=25)

        self.file_scrollbar = tk.Scrollbar(root, command=self.file_listbox.yview)
        self.sheet_scrollbar = tk.Scrollbar(root, command=self.sheet_listbox.yview)

        self.file_listbox.pack(side=tk.LEFT, padx=10, pady=10)
        self.sheet_listbox.pack(side=tk.LEFT, padx=10, pady=10)

        self.file_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.sheet_scrollbar.pack(side=tk.LEFT, fill=tk.Y)

        self.file_listbox.config(yscrollcommand=self.file_scrollbar.set)
        self.sheet_listbox.config(yscrollcommand=self.sheet_scrollbar.set)

        self.add_file_button = tk.Button(root, text="添加Excel文件", command=self.add_file, width=14)
        self.add_file_button.place(x=935, y=130)

        self.merge_workbooks_button = tk.Button(root, text="合并工作簿", command=self.merge_workbooks, width=14)
        self.merge_workbooks_button.place(x=935, y=200)

        self.merge_sheets_button = tk.Button(root, text="合并工作表", command=self.merge_sheets, width=14)
        self.merge_sheets_button.place(x=935, y=270)

        self.select_all_files_var = tk.IntVar(value=0)
        self.select_all_files_checkbox = tk.Checkbutton(root, text="全选文件", variable=self.select_all_files_var,
                                                        command=self.select_all_files)
        self.select_all_files_checkbox.place(x=50, y=10)

        self.select_all_sheets_var = tk.IntVar(value=0)
        self.select_all_sheets_checkbox = tk.Checkbutton(root, text="全选工作表", variable=self.select_all_sheets_var,
                                                         command=self.select_all_sheets)
        self.select_all_sheets_checkbox.place(x=480, y=10)

    def add_file(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        failed_files = []  # 存储解析失败的文件信息

        for file_path in file_paths:
            try:
                self.selected_files.append(file_path)
                self.file_listbox.insert(tk.END, file_path)
                self.update_sheet_listbox()
            except Exception as e:
                failed_files.append((file_path, str(e)))  # 存储失败文件信息

        if failed_files:
            self.show_failed_files_message(failed_files)

    def show_failed_files_message(self, failed_files):
        message = "以下文件无法解析：\n\n"
        for file_path, error_message in failed_files:
            message += f"{os.path.basename(file_path)} - {error_message}\n"

        messagebox.showerror("错误", message, parent=self.root)

    def update_sheet_listbox(self):
        self.sheet_listbox.delete(0, tk.END)
        self.selected_sheets = []

        for file_path in self.selected_files:
            workbook = pd.ExcelFile(file_path)
            for sheet_name in workbook.sheet_names:
                sheet_info = (file_path, sheet_name)  # Store tuple of (工作簿, 工作表)
                self.sheet_listbox.insert(tk.END, sheet_info)

    def select_all_files(self):
        selected_value = self.select_all_files_var.get()
        if selected_value == 1:
            self.file_listbox.select_set(0, tk.END)
        else:
            self.file_listbox.selection_clear(0, tk.END)

    def select_all_sheets(self):
        selected_value = self.select_all_sheets_var.get()
        if selected_value == 1:
            self.sheet_listbox.select_set(0, tk.END)
        else:
            self.sheet_listbox.selection_clear(0, tk.END)

    def merge_workbooks(self):
        if len(self.selected_files) >= 2:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            if save_path:
                with pd.ExcelWriter(save_path, engine="xlsxwriter") as merged_workbook:
                    for file_path in self.selected_files:
                        workbook = pd.ExcelFile(file_path)
                        for sheet_name in workbook.sheet_names:
                            new_sheet_name = self.get_unique_sheet_name(merged_workbook, sheet_name,
                                                                        os.path.basename(file_path))
                            df = pd.read_excel(file_path, sheet_name)
                            if not df.empty:
                                self.copy_df_to_excel(merged_workbook, df, new_sheet_name)

                self.update_sheet_listbox()

                messagebox.showinfo("Success", "Workbooks merged!")

    def get_unique_sheet_name(self, workbook, sheet_name, file_basename):
        if sheet_name in workbook.sheets:
            return f"{file_basename}_{sheet_name}"
        return sheet_name

    def copy_df_to_excel(self, writer, df, sheet_name):
        writer.sheets = dict((ws.title, ws) for ws in writer.sheets)
        if sheet_name in writer.sheets:
            return
        df.to_excel(writer, sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        for idx, column in enumerate(df):
            series = df[column]
            max_len = max((
                series.astype(str).map(len).max(),
                len(str(column))
            )) + 1
            worksheet.set_column(idx, idx, max_len)

    def merge_sheets(self):
        self.selected_sheets = [self.sheet_listbox.get(idx) for idx in self.sheet_listbox.curselection()]

        if len(self.selected_files) >= 1 and len(self.selected_sheets) >= 2:
            selected_sheets_data = []

            titles = None  # 用于存储标题行
            for file_path, sheet_name in self.selected_sheets:
                workbook = pd.ExcelFile(file_path)
                if sheet_name in workbook.sheet_names:
                    try:
                        df = pd.read_excel(file_path, sheet_name)
                        if titles is None:
                            titles = list(df.columns)
                        elif list(df.columns) != titles:  # 检查标题行是否相同
                            messagebox.showerror("Error", "Selected sheets have different titles!")
                            return
                        selected_sheets_data.append(df)
                    except Exception as e:
                        messagebox.showerror("Error",
                                             f"An error occurred while reading sheet '{sheet_name}' in file '{file_path}': {str(e)}")

            if selected_sheets_data:
                try:
                    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
                    if save_path:
                        merged_df = pd.concat(selected_sheets_data, axis=0)
                        merged_df.to_excel(save_path, index=False)
                        messagebox.showinfo("Success", "Sheets merged!")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred while merging sheets: {str(e)}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    app.run()