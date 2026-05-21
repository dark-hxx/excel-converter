import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk

import pandas as pd

from bank_converter import open_bank_converter_window, open_batch_converter_window
from utils import center_window

# 解决打包后路径问题
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.getcwd()

# 通用模板映射：历史文件
HISTORY_TEMPLATES_FILE = os.path.join(base_dir, 'history_templates.json')
HISTORY_MAPPINGS_FILE = os.path.join(base_dir, 'history_mappings.json')


# 全局变量（通用模板映射用）
save_dir = None
template_df = None
column_mapping = None


# ====================== 通用模板映射（旧版） ======================

def load_history_templates():
    try:
        with open(HISTORY_TEMPLATES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []


def save_history_template(template_path):
    history_templates = load_history_templates()
    if template_path not in history_templates:
        history_templates.append(template_path)
    with open(HISTORY_TEMPLATES_FILE, 'w', encoding='utf-8') as f:
        json.dump(history_templates, f, ensure_ascii=False, indent=4)


def load_history_mappings():
    try:
        with open(HISTORY_MAPPINGS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if isinstance(data, list):
                return {f'mapping_{i}': item for i, item in enumerate(data)}
            return data
    except FileNotFoundError:
        return {}


def save_history_mapping(mapping):
    history_mappings = load_history_mappings()
    history_mappings.append(mapping)
    with open(HISTORY_MAPPINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(history_mappings, f, ensure_ascii=False, indent=4)


def select_template():
    global template_df
    history_templates = load_history_templates()
    template_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("历史模板", [os.path.basename(t) for t in history_templates])])
    if template_path:
        try:
            template_df = pd.read_excel(template_path, sheet_name=0, dtype=str)
            log_text.insert(tk.END, f"模板文件 {template_path} 已成功加载\n")
            save_history_template(template_path)
        except Exception as e:
            messagebox.showerror("错误", f"加载模板文件时出错: {str(e)}")
            log_text.insert(tk.END, f"加载模板文件 {template_path} 出错: {str(e)}\n")


def set_column_mapping():
    global column_mapping
    if 'template_df' not in globals() or template_df is None:
        messagebox.showerror("错误", "请先选择模板文件")
        log_text.insert(tk.END, "设置列映射失败：请先选择模板文件\n")
        return

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, sheet_name=0, dtype=str)
        file_columns = df.columns.tolist()
    except Exception as e:
        messagebox.showerror("错误", f"读取文件 {file_path} 时出错: {str(e)}")
        log_text.insert(tk.END, f"读取文件 {file_path} 出错: {str(e)}\n")
        return

    mapping_window = tk.Toplevel(root)
    mapping_window.title("设置列映射")

    date_formats = [
        "", "yyyyMMdd", "yyyy/MM/dd", "yyyy-MM-dd", "yyyy年MM月dd日",
        "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd HH:mm:ss",
        "yyyy/MM/dd HH:mm:ss", "yyyyMMddHHmmss",
    ]

    template_columns = list(template_df.columns)
    tk.Label(mapping_window, text="模板列").grid(row=0, column=0)
    tk.Label(mapping_window, text="文件列").grid(row=0, column=1)
    tk.Label(mapping_window, text="分隔符").grid(row=0, column=2)
    tk.Label(mapping_window, text="输入时间格式").grid(row=0, column=3)
    tk.Label(mapping_window, text="输出时间格式").grid(row=0, column=4)

    combo_boxes = []
    entry_boxes = []
    input_date_combos = []
    output_date_combos = []
    for i, col in enumerate(template_columns):
        tk.Label(mapping_window, text=col).grid(row=i + 1, column=0)
        combo = ttk.Combobox(mapping_window, values=file_columns)
        combo.grid(row=i + 1, column=1)
        combo_boxes.append(combo)

        entry = tk.Entry(mapping_window)
        entry.grid(row=i + 1, column=2)
        entry_boxes.append(entry)

        input_date_combo = ttk.Combobox(mapping_window, values=date_formats, width=15)
        input_date_combo.grid(row=i + 1, column=3)
        input_date_combos.append(input_date_combo)

        output_date_combo = ttk.Combobox(mapping_window, values=date_formats, width=15)
        output_date_combo.grid(row=i + 1, column=4)
        output_date_combos.append(output_date_combo)

    def save_mapping():
        global column_mapping
        name = simpledialog.askstring("输入映射名称", "请输入当前映射关系的名称：")
        if not name:
            messagebox.showwarning("警告", "未输入映射名称，映射未保存")
            return
        column_mapping = {}
        split_info = {}
        date_format_info = {}
        for i, col in enumerate(template_columns):
            selected_col = combo_boxes[i].get()
            split_symbol = entry_boxes[i].get()
            input_fmt = input_date_combos[i].get()
            output_fmt = output_date_combos[i].get()
            if selected_col:
                column_mapping[col] = selected_col
                if split_symbol:
                    split_info[selected_col] = split_symbol
                if input_fmt and output_fmt:
                    date_format_info[col] = {"input": input_fmt, "output": output_fmt}
        column_mapping['split_info'] = split_info
        column_mapping['date_format_info'] = date_format_info
        mapping_window.destroy()
        log_text.insert(tk.END, "列映射已成功保存\n")
        history_mappings = load_history_mappings()
        history_mappings[name] = column_mapping
        with open(HISTORY_MAPPINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(history_mappings, f, ensure_ascii=False, indent=4)

    save_button = tk.Button(mapping_window, text="保存映射", command=save_mapping)
    save_button.grid(row=len(template_columns) + 1, columnspan=5)


def select_save_directory():
    global save_dir
    save_dir = filedialog.askdirectory()
    if save_dir:
        log_text.insert(tk.END, f"保存目录已设置为: {save_dir}\n")


def convert_excel_files():
    if 'template_df' not in globals() or 'column_mapping' not in globals() or column_mapping is None:
        messagebox.showerror("错误", "请先选择模板并设置列映射")
        return
    if save_dir is None:
        messagebox.showerror("错误", "请先选择保存目录")
        return

    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_paths:
        return

    split_info = column_mapping.get('split_info', {})
    date_format_info = column_mapping.get('date_format_info', {})
    template_to_file_mapping = {k: v for k, v in column_mapping.items() if k not in ('split_info', 'date_format_info')}

    for file_path in file_paths:
        try:
            sheet_names = pd.ExcelFile(file_path).sheet_names

            for sheet_name in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)

                log_text.insert(tk.END, f"调试信息: 文件 {file_path} 工作表 {sheet_name} - 开始处理\n")
                log_text.insert(tk.END, f"调试信息: column_mapping: {column_mapping}\n")

                file_columns_ordered = [template_to_file_mapping[k] for k in template_to_file_mapping.keys()]
                missing_columns = [col for col in file_columns_ordered if col not in df.columns]
                if missing_columns:
                    log_text.insert(tk.END, f"错误: 文件 {file_path} 工作表 {sheet_name} 缺少列: {missing_columns}\n")
                    continue

                mapped_df = df[file_columns_ordered].copy()
                template_columns_ordered = list(template_to_file_mapping.keys())
                mapped_df.columns = template_columns_ordered

                reverse_column_mapping = {v.strip().lower(): k.strip().lower() for k, v in template_to_file_mapping.items()}
                adjusted_split_info = {}
                for file_col, split_symbol in split_info.items():
                    template_col_lower = reverse_column_mapping.get(file_col.strip().lower())
                    if template_col_lower:
                        for orig_col in template_columns_ordered:
                            if orig_col.strip().lower() == template_col_lower:
                                adjusted_split_info[orig_col] = split_symbol
                                break
                    else:
                        log_text.insert(tk.END, f"警告: split_info 中的列 '{file_col}' 未在映射中找到\n")

                if adjusted_split_info:
                    new_rows = []
                    for index, row in mapped_df.iterrows():
                        try:
                            split_values_dict = {}
                            max_length = 1
                            for col_name, split_symbol in adjusted_split_info.items():
                                if col_name in row.index:
                                    if pd.notna(row[col_name]):
                                        values = str(row[col_name]).split(split_symbol)
                                        values = [v.strip().strip(split_symbol).strip() for v in values]
                                        values = [v for v in values if v]
                                        if values:
                                            split_values_dict[col_name] = values
                                            max_length = max(max_length, len(values))
                                        else:
                                            split_values_dict[col_name] = [str(row[col_name])]
                                    else:
                                        split_values_dict[col_name] = ['']
                            if split_values_dict:
                                for i in range(max_length):
                                    new_row = row.copy()
                                    for col_name, values in split_values_dict.items():
                                        new_row[col_name] = values[i] if i < len(values) else ''
                                    new_rows.append(new_row)
                                log_text.insert(tk.END, f"信息: 行 {index} 拆分为 {max_length} 行\n")
                            else:
                                new_rows.append(row)
                        except Exception as row_error:
                            log_text.insert(tk.END, f"错误: 处理行 {index} 时出错: {str(row_error)}\n")
                            new_rows.append(row)
                            continue
                    mapped_df = pd.DataFrame(new_rows)

                if date_format_info:
                    from datetime import datetime

                    def convert_date_format(fmt):
                        if not fmt:
                            return fmt
                        replacements = [
                            ('yyyy', '%Y'), ('yy', '%y'), ('MM', '%m'),
                            ('dd', '%d'), ('HH', '%H'), ('mm', '%M'), ('ss', '%S'),
                        ]
                        result = fmt
                        for old, new in replacements:
                            result = result.replace(old, new)
                        return result

                    for col, fmt_config in date_format_info.items():
                        if col in mapped_df.columns:
                            input_fmt = convert_date_format(fmt_config.get('input', ''))
                            output_fmt = convert_date_format(fmt_config.get('output', ''))
                            if input_fmt and output_fmt:
                                def convert_date(val, in_fmt=input_fmt, out_fmt=output_fmt):
                                    if pd.isna(val) or str(val).strip() == '':
                                        return ''
                                    try:
                                        dt = datetime.strptime(str(val).strip(), in_fmt)
                                        return dt.strftime(out_fmt)
                                    except ValueError:
                                        return str(val)
                                mapped_df[col] = mapped_df[col].apply(convert_date)

                all_template_columns = list(template_df.columns)
                for col in all_template_columns:
                    if col not in mapped_df.columns:
                        mapped_df[col] = ''
                mapped_df = mapped_df[all_template_columns]

                file_name = os.path.splitext(os.path.basename(file_path))[0]
                if len(sheet_names) > 1:
                    file_name += f'_{sheet_name}'
                save_path = os.path.join(save_dir, f'{file_name}.xlsx')

                for col in mapped_df.columns:
                    mapped_df[col] = mapped_df[col].astype(str)
                    mapped_df[col] = mapped_df[col].replace('nan', '')

                try:
                    from openpyxl import load_workbook
                    from openpyxl.styles import numbers
                    mapped_df.to_excel(save_path, index=False, sheet_name='Sheet1', engine='openpyxl')
                    wb = load_workbook(save_path)
                    ws = wb.active
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                        for cell in row:
                            cell.number_format = numbers.FORMAT_TEXT
                    wb.save(save_path)
                    wb.close()
                    log_text.insert(tk.END, f"成功: 文件 {file_path} 工作表 {sheet_name} 转换完成，保存为 {save_path}\n")
                except PermissionError:
                    error_msg = f"无法保存文件 {save_path}\n可能原因：\n1. 文件正在被其他程序（如Excel）打开\n2. 没有写入权限\n请关闭文件后重试"
                    log_text.insert(tk.END, f"错误: {error_msg}\n")
                    messagebox.showerror("文件保存失败", error_msg)
                    continue
                except Exception as save_error:
                    log_text.insert(tk.END, f"错误: 保存文件时出错: {str(save_error)}\n")
                    messagebox.showerror("保存失败", f"保存文件时出错: {str(save_error)}")
                    continue

        except Exception as e:
            log_text.insert(tk.END, f"错误: 处理文件 {file_path} 时出错，详细错误信息: {str(e)}\n")
            messagebox.showerror("错误", f"处理文件 {file_path} 时出错: {str(e)}")

    messagebox.showinfo("完成", "所有文件转换完成！")


def show_history_templates():
    history_templates = load_history_templates()
    if not history_templates:
        messagebox.showinfo("历史模板", "暂无历史模板记录")
    else:
        template_window = tk.Toplevel(root)
        template_window.title("历史模板")
        for template_path in history_templates:
            tk.Button(template_window, text=template_path, command=lambda path=template_path: use_history_template(path)).pack(pady=5)


def use_history_template(template_path):
    global template_df
    try:
        template_df = pd.read_excel(template_path, sheet_name=0, dtype=str)
        log_text.insert(tk.END, f"历史模板文件 {template_path} 已成功加载\n")
        save_history_template(template_path)
    except Exception as e:
        messagebox.showerror("错误", f"加载历史模板文件时出错: {str(e)}")
        log_text.insert(tk.END, f"加载历史模板文件 {template_path} 出错: {str(e)}\n")


def show_history_mappings():
    history_mappings = load_history_mappings()
    if not history_mappings:
        messagebox.showinfo("历史映射", "暂无历史映射记录")
    else:
        mapping_window = tk.Toplevel(root)
        mapping_window.title("历史映射")
        for name in history_mappings:
            tk.Button(mapping_window, text=name, command=lambda n=name: use_history_mapping(history_mappings[n])).pack(pady=5)


def use_history_mapping(mapping):
    global column_mapping
    column_mapping = mapping
    log_text.insert(tk.END, "历史映射已成功应用\n")


# ====================== 主窗口 ======================

root = tk.Tk()
root.title("Excel 文件转换器")
center_window(root, 520, 680)

# ---------------- 银行流水转换分组 ----------------
bank_frame = ttk.LabelFrame(root, text=' 银行流水转换（推荐） ', padding=10)
bank_frame.pack(fill='x', padx=15, pady=(15, 8))

bank_btn_row = tk.Frame(bank_frame)
bank_btn_row.pack(fill='x')

tk.Button(bank_btn_row, text='单文件转换',
          command=lambda: open_bank_converter_window(root),
          bg='#2196F3', fg='white', height=2
          ).pack(side='left', expand=True, fill='x', padx=2)
tk.Button(bank_btn_row, text='批量混合转换',
          command=lambda: open_batch_converter_window(root),
          bg='#1565C0', fg='white', height=2
          ).pack(side='left', expand=True, fill='x', padx=2)

tk.Label(bank_frame, text='内置 12 家银行规则；批量模式可逐文件指定银行并合并到一个文件',
         fg='#666').pack(pady=(6, 0))

# ---------------- 通用模板映射分组 ----------------
generic_frame = ttk.LabelFrame(root, text=' 通用模板映射（自定义） ', padding=10)
generic_frame.pack(fill='x', padx=15, pady=8)

row1 = tk.Frame(generic_frame)
row1.pack(fill='x', pady=2)
tk.Button(row1, text='选择模板文件', command=select_template).pack(side='left', expand=True, fill='x', padx=2)
tk.Button(row1, text='设置列映射', command=set_column_mapping).pack(side='left', expand=True, fill='x', padx=2)

row2 = tk.Frame(generic_frame)
row2.pack(fill='x', pady=2)
tk.Button(row2, text='选择生成文件路径', command=select_save_directory).pack(side='left', expand=True, fill='x', padx=2)
tk.Button(row2, text='转换 Excel 文件', command=convert_excel_files,
          bg='#4CAF50', fg='white').pack(side='left', expand=True, fill='x', padx=2)

row3 = tk.Frame(generic_frame)
row3.pack(fill='x', pady=2)
tk.Button(row3, text='显示历史模板', command=show_history_templates).pack(side='left', expand=True, fill='x', padx=2)
tk.Button(row3, text='显示历史映射', command=show_history_mappings).pack(side='left', expand=True, fill='x', padx=2)

# ---------------- 日志区 ----------------
log_frame = ttk.LabelFrame(root, text=' 运行日志 ', padding=5)
log_frame.pack(fill='both', expand=True, padx=15, pady=(8, 15))

log_text = scrolledtext.ScrolledText(log_frame, height=8)
log_text.pack(fill='both', expand=True)

root.mainloop()
