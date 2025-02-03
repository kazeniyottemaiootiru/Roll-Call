import os
import pickle
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


class Excelsetting:
    def __init__(self, root):
        self.root = root
        self.root.title("点名器名单生成器")
        self.root.configure(bg = "white")

        # 显示文本
        self.text_show = tk.Label(root, text = "欢迎使用本点名器生成器", font = "黑体 24", bg = "white")
        self.text_show.pack(expand = True)

        self.text_note = tk.Label(root, text = "打包过程可能会出现无响应的情况，请耐心等待。", font = "黑体 16",
                                  bg = "white")
        self.text_note.pack(expand = True)

        frame_top = tk.Frame(root, bg = "white")
        frame_top.pack(pady = 20)

        self.btn_select_file = tk.Button(frame_top, text = "选择Excel文件", command = self.load_excel,
                                         font = "微软雅黑 14")
        self.btn_select_file.pack(side = "left", padx = 10)

        # Sheet 选择框
        self.sheet_var = tk.StringVar()
        self.sheet_dropdown = tk.OptionMenu(frame_top, self.sheet_var, "")
        self.sheet_dropdown.pack(pady = 10)

        # 列名选择框
        self.column_var = tk.StringVar()
        self.column_dropdown = tk.OptionMenu(frame_top, self.column_var, "")
        self.column_dropdown.pack(pady = 10)

        frame_bottom = tk.Frame(root, bg = "white")
        frame_bottom.pack(anchor = "e", padx = 20)

        self.btn_confirm_column = tk.Button(frame_bottom, text = "打包", command = self.generate_exe,
                                            state = tk.DISABLED, font = "微软雅黑 12")
        self.btn_confirm_column.pack(side = "left", padx = 10)

        self.btn_cancel = tk.Button(frame_bottom, text = "取消", command = self.root.quit, font = "微软雅黑 12")
        self.btn_cancel.pack(side = "left", padx = 10)

        self.df = None
        self.file_path = ""

    def load_excel(self):
        self.file_path = filedialog.askopenfilename(filetypes = [("Excel文件", "*.xlsx;*.xls")])
        if not self.file_path:
            return

        try:
            excel_file = pd.ExcelFile(self.file_path)
            sheets = excel_file.sheet_names  # 获取所有 Sheet 名称

            if not sheets:
                messagebox.showerror("错误", "Excel文件中没有可用的Sheet！")
                return

            # 先清空原有的 Sheet 列表
            self.sheet_dropdown["menu"].delete(0, "end")

            # 逐个添加 Sheet 名称
            for sheet in sheets:
                self.sheet_dropdown["menu"].add_command(label = sheet,
                                                        command = lambda value = sheet: self.update_columns(value))

            # 默认选择第一个 Sheet，并更新列名
            self.sheet_var.set(sheets[0])
            self.update_columns(sheets[0])

            messagebox.showinfo("成功", "Excel文件加载成功，请选择Sheet！")

        except Exception as e:
            messagebox.showerror("错误", f"无法读取Excel文件:\n{e}")

    def update_columns(self, sheet_name):
        """当用户选择 Sheet 时，更新列名"""
        try:
            self.sheet_var.set(sheet_name)  # 这句设置变量值
            self.sheet_dropdown.update_idletasks()  # 强制 UI 更新

            self.df = pd.read_excel(self.file_path, sheet_name = sheet_name)
            columns = self.df.columns.tolist()

            if not columns:
                messagebox.showerror("错误", "所选 Sheet 没有可用列！")
                return

            # 先清空原有的列名列表
            self.column_dropdown["menu"].delete(0, "end")

            # 逐个添加列名
            for col in columns:
                self.column_dropdown["menu"].add_command(label = col,
                                                         command = lambda value = col: self.column_var.set(value))

            # 默认选择第一列
            self.column_var.set(columns[0])

            # 启用打包按钮
            self.btn_confirm_column.config(state = tk.NORMAL)
            messagebox.showinfo("成功", f"已选择 Sheet: {sheet_name}，请继续选择姓名列！")
        except Exception as e:
            messagebox.showerror("错误", f"无法读取所选 Sheet:\n{e}")

    def update_selected_column(self, column_name):
        """更新选择的列，并刷新 OptionMenu 显示"""
        self.column_var.set(column_name)  # 更新 OptionMenu 变量
        self.column_dropdown.update_idletasks()  # 刷新 UI

    def generate_exe(self):
        column_name = self.column_var.get()
        if not column_name:
            messagebox.showerror("错误", "请选择一个列！")
            return

        try:
            names = self.df[column_name].dropna().astype(str).tolist()
            with open("names.pkl", "wb") as f:
                pickle.dump(names, f)

            self.package_with_pyinstaller()

            messagebox.showinfo("成功", "点名器已生成，请到 'dist/' 下找 'Rollcall.exe' 文件。")
        except Exception as e:
            messagebox.showerror("错误", f"生成 exe 时出错:\n{e}")

    def package_with_pyinstaller(self):
        current_directory = os.getcwd()

        pyinstaller_command = [
            "pyinstaller",
            "--onefile",
            "--noconsole",
            "--add-data", "names.pkl;.",
            "Rollcall.py"
        ]

        try:
            subprocess.check_call(pyinstaller_command, cwd = current_directory)
        except subprocess.CalledProcessError as e:
            print(f"打包失败: {e}")
            sys.exit(1)


if __name__ == "__main__":
    root = tk.Tk()
    app = Excelsetting(root)
    root.mainloop()
