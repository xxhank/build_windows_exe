#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re
import textwrap
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd


class FlowFrame(ttk.Frame):
  """模拟流式布局，按钮从左到右排列，超过宽度自动换行"""

  def __init__(self, parent, padding=5, **kwargs):
    super().__init__(parent, **kwargs)
    self.padding = padding
    self.widgets = []

  def add_widget(self, widget):
    self.widgets.append(widget)
    self._reflow()

  def _reflow(self):
    # 清空所有布局
    for w in self.widgets:
      w.grid_forget()

    x = y = 0
    max_width = self.winfo_width() or self.winfo_reqwidth()
    row = 0
    col = 0
    for w in self.widgets:
      w.update_idletasks()
      w_width = w.winfo_reqwidth()
      # 换行
      if x + w_width > max_width:
        row += 1
        x = 0
        col = 0
      w.grid(row=row, column=col, padx=self.padding, pady=self.padding, sticky='w')
      x += w_width + self.padding
      col += 1

class ExcelToolApp:
  def __init__(self, root):
    self.root = root
    self.root.title("Excel工具")
    self.center_window(self.root, 600, 300)
    self.root.resizable(True, False)

    # 状态变量
    self.excel_path = tk.StringVar()
    self.output_name = tk.StringVar()
    self.sheet_name = tk.StringVar()

    # UI 结构
    self._build_ui()
    # 监听输入变化

    self.output_name.trace_add("write", lambda v1, v2, v3: self.update_execute_button())

  def center_window(self, window, width=600, height=400):
    # 获取屏幕宽高
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # 计算居中的坐标
    x = int((screen_width - width) / 2)
    y = int((screen_height - height) / 2)

    # 设置窗口大小和位置
    window.geometry(f"{width}x{height}+{x}+{y}")

  def _build_ui(self):
    # ---------- 顶部说明 ----------
    self.top_frame = ttk.Frame(self.root, padding=10)
    self.top_frame.pack(fill='x', side='top')

    # 滚动条
    scrollbar = ttk.Scrollbar(self.top_frame, orient='vertical')
    scrollbar.pack(side='right', fill='y')

    # Text 控件
    info_text = tk.Text(self.top_frame, height=3, wrap='word', yscrollcommand=scrollbar.set)
    info_text.pack(fill='x', expand=True)
    info_text.insert('1.0', textwrap.dedent("""
      各种工具集合,通过底部的按钮使用
    """).strip())
    info_text.configure(state='disabled')  # 设置为只读

    scrollbar.config(command=info_text.yview)

    frame = ttk.Frame(self.root, padding=20)
    frame.pack(fill='both', expand=True)

    frame.columnconfigure(0, weight=0)  # 第一列固定
    frame.columnconfigure(1, weight=1)  # 中间列可拉伸
    frame.columnconfigure(2, weight=0)  # 最后一列固定

    # 1️⃣ Excel 文件选择
    ttk.Label(frame, text="选择Excel文件:").grid(row=0, column=0, sticky="w")
    ttk.Entry(frame, textvariable=self.excel_path, width=45).grid(row=0, column=1, padx=5, sticky="ew")
    ttk.Button(frame, text="浏览", command=self.choose_excel).grid(row=0, column=2, sticky="w")

    # 2️⃣ Sheet 名称选择
    ttk.Label(frame, text="选择Sheet:").grid(row=1, column=0, sticky="w", pady=10)
    self.sheet_box = ttk.Combobox(frame, textvariable=self.sheet_name, state="disabled", width=42)
    self.sheet_box.grid(row=1, column=1, padx=5, sticky="ew")
    self.sheet_box.bind("<<ComboboxSelected>>", lambda e: self.update_execute_button())

    # 3️⃣ 输出文件名输入
    ttk.Label(frame, text="输出文件名:").grid(row=2, column=0, sticky="w", pady=10)
    ttk.Entry(frame, textvariable=self.output_name, width=45).grid(row=2, column=1, padx=5, sticky="ew")
    # ttk.Label(frame, text=".xlsx").grid(row=2, column=2, sticky="w")

    self.bottom_frame = FlowFrame(frame, padding=5)
    self.bottom_frame.grid(row=4, column=0, columnspan=3, sticky='ew', pady=10)

    self.execute_btn = ttk.Button(self.bottom_frame, text="生成 select xx union all", command=lambda :self.run(self.generate_select_union_all), state="disabled")
    # self.execute_btn.grid(row=4, column=1, pady=30)
    self.bottom_frame.add_widget(self.execute_btn)


  def choose_excel(self):
    path = filedialog.askopenfilename(
      title="选择Excel文件",
      filetypes=[("Excel 文件", "*.xlsx *.xls")]
    )
    if not path:
      return

    self.excel_path.set(path)
    try:
      xls = pd.ExcelFile(path)
      sheets = xls.sheet_names
      self.sheet_box["values"] = sheets
      self.sheet_box["state"] = "readonly"

      self.sheet_box.current(0)  # ✅ 默认选中第一个 sheet
      self.sheet_name.set(sheets[0])  # ✅ 同步变量值

      # 自动设置默认输出名：原文件名 + ".out.xlsx"
      src = Path(path)
      default_output = f"{src.stem}.out"
      self.output_name.set(default_output)

      # messagebox.showinfo("提示", f"已读取到 {len(sheets)} 个 Sheet。")
    except Exception as e:
      messagebox.showerror("错误", f"无法读取Excel文件：{e}")

    self.update_execute_button()

  def update_execute_button(self):
    """
    判断是否所有输入有效:
    1. 已选择Excel文件
    2. 已选择Sheet
    3. 输出文件名合法 (必须含扩展名 .xlsx 或 .xls)
    """
    excel_ok = bool(self.excel_path.get())
    sheet_ok = bool(self.sheet_name.get())

    # 检查输出文件名合法性
    output = self.output_name.get().strip()
    valid_name = re.match(r"^[^\\/:*?\"<>|\r\n]+\..+?$", output) is not None

    if excel_ok and sheet_ok and valid_name:
      self.execute_btn["state"] = "normal"
    else:
      self.execute_btn["state"] = "disabled"

  def run(self, process:Callable[[Path,str,Path],None]):
    try:
      excel = Path(self.excel_path.get())
      output_path = excel.parent / self.output_name.get().strip()
      sheet = self.sheet_name.get()

      # df = pd.read_excel(excel, sheet_name=sheet)
      # ✨ 示例操作：简单写回输出文件
      # df.to_excel(output_path, index=False)
      process(excel, sheet, output_path)

      messagebox.showinfo("完成", f"✅ 处理完成！结果已保存到:\n{output_path}")
    except Exception as e:
      messagebox.showerror("错误", str(e))

  def generate_select_union_all(self,excel:Path, sheet_name:str, out_path:Path):
    df = pd.read_excel(excel, sheet_name)
    col_name = df.columns[0]
    values = df[col_name].dropna().astype(str).tolist()
    result:list[str] = []
    for val in values:  # 第一列的数据
      result.append(f'select \'{val}\'')
    result_all_text = ' union all\n'.join(result)
    out_path.write_text(result_all_text)


if __name__ == "__main__":
  root = tk.Tk()
  app = ExcelToolApp(root)
  root.mainloop()
