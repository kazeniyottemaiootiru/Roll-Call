import pickle
import random
import tkinter as tk
from tkinter import messagebox
import sys
import os


class RollCallWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("点名器")
        self.root.geometry("800x400")

        try:
            # 支持 PyInstaller 打包后的路径
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.abspath(".")

            # 加载嵌入的 names.pkl
            names_path = os.path.join(base_path, "names.pkl")
            with open(names_path, "rb") as f:
                self.names = pickle.load(f)

        except Exception as e:
            messagebox.showerror("错误", f"无法加载名单文件: {e}")
            self.names = []

        self.text_show = tk.Label(self.root, text="点击按钮开始点名", font="微软雅黑 50")
        self.text_show.pack(pady=30)

        self.btn_roll_call = tk.Button(self.root, text="开始点名", command=self.random_name, font="微软雅黑 16")
        self.btn_roll_call.pack(pady=20)

        self.btn_exit = tk.Button(self.root, text="退出", command=self.root.quit, font="微软雅黑 16")
        self.btn_exit.pack(pady=10)

        self.root.mainloop()

    def random_name(self):
        if not self.names:
            messagebox.showinfo("提示", "名单为空！")
            return
        selected_name = random.choice(self.names)
        self.text_show.config(text=f"点名：{selected_name}")


if __name__ == "__main__":
    RollCallWindow()