import os
from tkinter import Tk, Entry, Button, Label, filedialog, messagebox, StringVar, Menu, Toplevel
from typing import List

from globals import Global
from scanners import ExcelScanner, WordScanner


class Window:
    template_path: str = None
    table_path: str = None
    sheets_list_strs: List[str] = []

    def __init__(self, wname: str, width: int, height: int):
        self._window = Tk()
        self._window.title(wname)
        self._window.geometry(f'{width}x{height}')

        def template_btn_click():
            self.template_path = filedialog.askopenfilename()
            _txt = StringVar()
            _txt.set(self.template_path)
            self.template_entry.config(textvariable=_txt)

        self.template_lbl = Label(self._window, text=u"WORD模板：")
        self.template_lbl.place(x=20, y=30)
        self.template_entry = Entry(self._window, width=35)
        self.template_entry.place(x=20, y=50, height=25)
        self.template_btn = Button(self._window, text=u"浏览", command=template_btn_click)
        self.template_btn.place(x=270, y=50, height=25)

        def table_btn_click():
            self.table_path = filedialog.askopenfilename()
            _txt = StringVar()
            _txt.set(self.table_path)
            self.table_entry.config(textvariable=_txt)

        self.table_lbl = Label(self._window, text=u"EXCEL表格")
        self.table_lbl.place(x=20, y=80)
        self.table_entry = Entry(self._window, width=35)
        self.table_entry.place(x=20, y=100, height=25)
        self.table_btn = Button(self._window, text=u"浏览", command=table_btn_click)
        self.table_btn.place(x=270, y=100, height=25)

        self.sheets_lbl = Label(self._window, text=u"填写表单")
        self.sheets_lbl.place(x=20, y=130)
        self.sheets_list = Entry(self._window, width=35)
        self.sheets_list.place(x=20, y=150, height=25)

        self.generate_btn = Button(self._window, text=u"生成")
        self.generate_btn.place(x=350, y=50, width=80, height=120)

        self.main_menu = Menu(self._window)
        self.opt_menu = Menu(self.main_menu, tearoff=False)
        self.opt_menu.add_command(label=u"配置...", command=self.config_command)
        self.main_menu.add_cascade(label=u"选项", menu=self.opt_menu)
        self.main_menu.add_command(label=u"帮助", command=self.help_command)
        self._window.config(menu=self.main_menu)

    def config_command(self):
        dialog = Toplevel(self._window)
        dialog.title(u"配置")
        dialog.geometry('500x200')
        dialog.grab_set()
        dialog.protocol('WM_DELETE_WINDOW', lambda: dialog.destroy())

        dst = Global.dst_path

        def dst_btn_click():
            dst = filedialog.askdirectory()
            _txt = StringVar()
            _txt.set(dst)
            dst_entry.config(textvariable=_txt)

        dst_lbl = Label(dialog, text=u"目标文件生成路径")
        dst_lbl.place(x=20, y=30)
        dst_entry = Entry(dialog)
        dst_entry.insert(0, dst)
        dst_entry.place(x=20, y=50, width=200, height=25)
        dst_btn = Button(dialog, text=u"浏览", command=dst_btn_click)
        dst_btn.place(x=220, y=50)

        para_font_lbl = Label(dialog, text=u"段落字号")
        para_font_lbl.place(x=20, y=80)
        para_font_entry = Entry(dialog)
        para_font_entry.insert(0, "16")
        para_font_entry.place(x=20, y=100, width=200, height=25)

        tab_font_lbl = Label(dialog, text=u"表格字号")
        tab_font_lbl.place(x=20, y=130)
        tab_font_entry = Entry(dialog)
        tab_font_entry.insert(0, "12")
        tab_font_entry.place(x=20, y=150, width=200, height=25)

        def ok_btn_click():
            Global.dst_path = dst
            Global.paragraph_font_size = int(para_font_entry.get()) if str(para_font_entry.get()).isdigit() else 16
            Global.table_font_size = int(tab_font_entry.get()) if str(tab_font_entry.get()).isdigit() else 12
            dialog.destroy()

        ok_btn = Button(dialog, text=u"确定", command=ok_btn_click)
        ok_btn.place(x=350, y=150, width=100)

    def help_command(self):
        messagebox.showinfo("提示", "使用说明详见https://github.com/littlebutt/abstract_generator")

    def bind_command(self):
        def check_and_bind():
            _v = self.sheets_list.get()
            if _v is None or _v.strip() == '':
                messagebox.showwarning(u"注意", u"请输入至少一个表单名，多个表单用半角逗号隔开")
            self.sheets_list_strs = _v.split(',')
            if self.template_path is None or self.table_path is None:
                messagebox.showwarning("注意", "请检查输入的路径是否正确")
                return
            Global.sheets = self.sheets_list_strs
            Global.excel_scanner = ExcelScanner(self.table_path)
            word = WordScanner(self.template_path, Global.dst_path + r"/result.docx")
            try:
                word.scan()
            except Exception as e:
                messagebox.showerror("错误", e)
                print(e.args)

        self.generate_btn.config(command=check_and_bind)

    def run(self):
        self._window.mainloop()
