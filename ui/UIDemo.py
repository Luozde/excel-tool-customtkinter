import os
import platform
import subprocess
import sys
import time
import tkinter
from tkinter import filedialog

import customtkinter


from core.DataTransformer import DataTransformer
from core.ExcelExporter import ExcelExporter


class App(customtkinter.CTk):
    choose_file = None
    export_dir = os.getcwd()
    export_file_name = None

    def __init__(self):
        super().__init__()

        self.title("订单配货表生成工具")
        self.geometry(f"{650}x{500}")
        self.grid_columnconfigure((0, 1), weight=1)
        self.grid_rowconfigure((5, 5), weight=1)

        # 选择订单文件
        self.file_btn = customtkinter.CTkButton(self, text="选择订单文件", width=25, command=self.file_button_callback)
        self.file_btn.grid(row=0, column=0, sticky="e", padx=20, pady=(20,0))
        self.file_path_label = customtkinter.CTkLabel(self, text="请选择订单文件")
        self.file_path_label.grid(row=0, column=1, columnspan=2, pady=(20, 0), sticky="w")

        # 设置导出文件夹
        self.export_dir_btn = customtkinter.CTkButton(self, text="设置导出文件夹", width=25, command=self.export_dir_btn_callback)
        self.export_dir_btn.grid(row=1, column=0, sticky="e", padx=20, pady=(20, 0))
        self.export_path_label = customtkinter.CTkLabel(self, text="默认当前目录 \n" + self.export_dir)
        self.export_path_label.grid(row=1, column=1, columnspan=2, pady=(20, 0), sticky="w")

        # 设置导出文件名
        self.export_dir_btn = customtkinter.CTkButton(self, text="设置导出文件名", width=25, state="disabled")
        self.export_dir_btn.grid(row=2, column=0, sticky="e", padx=20, pady=(20, 0))
        self.file_name_entry = customtkinter.CTkEntry(self, placeholder_text="默认：订单文件名-配货表-时间戳", width=300)
        self.file_name_entry.grid(row=2, column=1, columnspan=2, pady=(20, 0), sticky="nsw")

        # 执行按钮
        self.execute_btn = customtkinter.CTkButton(self, text="执行", command=self.execute_button_callback)
        self.execute_btn.grid(row=3,  column=1, pady=20, sticky="w")

        # 执行日志
        self.textbox_label = customtkinter.CTkLabel(self, text="处理日志", font=customtkinter.CTkFont(size=16, weight="bold"))
        self.textbox_label.grid(row=4, column=0, padx=20, pady=5, sticky="w")
        self.textbox = customtkinter.CTkTextbox(self)
        self.textbox.grid(row=5, columnspan=2, padx=20, pady=(0,20), sticky="nsew")

    # 选择文件按钮回调
    def file_button_callback(self):
        # 通过filedialog获取文件
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx'), ('Excel Files', '*.xls')])
        self.textbox.delete("0.0", "end")
        # 检查是否选择了文件
        if file_path:
            self.print_logs("选择文件：" + file_path + "\n", True)
            file_name = os.path.basename(file_path)
            self.choose_file = file_path
            directory_path = os.path.dirname(file_path)
            self.export_path_label.configure(text=directory_path)
            self.export_dir = directory_path
            # 获取时间戳
            if file_name:
                self.file_btn.configure(text="重新选择")
                self.file_name_entry.delete(0, "end")
                self.file_name_entry.insert(0, file_name.split(".")[0]+"-配货表-" + time.strftime("%m%d%H%M", time.localtime()))
            self.file_path_label.configure(text=file_name)

    # 执行按钮回调
    def execute_button_callback(self):
        # 检查是否选择了文件
        if self.choose_file:
            file_name = self.file_name_entry.get()
            # 如果file_name为空，使用文件名
            if file_name:
                file_name = file_name.split(".")[0]
            else:
                file_name = os.path.basename(self.choose_file).split(".")[0]+"-配货表-" + time.strftime("%m%d%H%M", time.localtime())
                self.file_name_entry.insert(0, file_name)
            self.textbox.delete("0.0", "end")
            self.print_logs("开始执行\n", True)
            data = DataTransformer.process_excel_file(self, self.choose_file)
            if not data:
                self.print_logs("执行完成\n", True)
                return
            summarize_sku_items = DataTransformer.summarize_sku_items(self, data)
            ExcelExporter.export_to_excel(self, summarize_sku_items, self.export_dir + "/" + file_name + ".xlsx")
            self.print_logs("输出订单汇总文件路径：" + self.export_dir + "/" + file_name + ".xlsx\n", True)
            self.print_logs("执行完成\n", True)
            self.execute_finish_confirmation()

    def export_dir_btn_callback(self):
        # 通过filedialog获取目录
        directory_path = filedialog.askdirectory()
        # 检查是否选择了文件
        if directory_path:
            self.export_path_label.configure(text=directory_path)
            self.export_dir = directory_path

    def print_logs(self, msg, end):
        if end:
            self.textbox.insert("end", str(msg) + "\n")
        else:
            self.textbox.insert("0.0", str(msg) + "\n")

    def open_dir_btn_callback(self):
        directory_path = self.export_dir
        system = platform.system()
        if system == 'Windows':
            directory_path = os.path.normpath(directory_path)
            subprocess.run(['explorer', directory_path], check=True)
        elif system == 'Darwin':
            subprocess.run(['open', directory_path], check=True)
        elif system == 'Linux':
            subprocess.run(['xdg-open', directory_path], check=True)
        else:
            print("当前操作系统不支持直接打开", system)
            sys.exit(1)

    def execute_finish_confirmation(self):
        result = tkinter.messagebox.askyesno("确认框", "执行完成，处理日志已输出到窗口，是否打开导出目录？")
        if result:
            print("执行操作")
            self.open_dir_btn_callback()
        else:
            print("取消操作")