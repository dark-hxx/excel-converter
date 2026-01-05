import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk

import pandas as pd

# 解决打包后路径问题
if getattr(sys, 'frozen', False):
    # 打包后路径
    base_dir = os.path.dirname(sys.executable)
else:
    # 开发环境路径
    base_dir = os.getcwd()

# 定义历史记录文件路径
HISTORY_TEMPLATES_FILE = os.path.join(base_dir, 'history_templates.json')
HISTORY_MAPPINGS_FILE = os.path.join(base_dir, 'history_mappings.json')


# 全局变量，用于存储保存目录
save_dir = None
template_df = None
column_mapping = None

# 加载历史模板
def load_history_templates():
    try:
        with open(HISTORY_TEMPLATES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []

# 保存历史模板
def save_history_template(template_path):
    history_templates = load_history_templates()
    if template_path not in history_templates:
        history_templates.append(template_path)
    with open(HISTORY_TEMPLATES_FILE, 'w', encoding='utf-8') as f:
        json.dump(history_templates, f, ensure_ascii=False, indent=4)

# 加载历史映射
def load_history_mappings():
    try:
        with open(HISTORY_MAPPINGS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # 如果加载的数据是列表，转换为字典
            if isinstance(data, list):
                new_data = {}
                for i, item in enumerate(data):
                    new_data[f'mapping_{i}'] = item
                return new_data
            return data
    except FileNotFoundError:
        return {}

# 保存历史映射
def save_history_mapping(mapping):
    history_mappings = load_history_mappings()
    history_mappings.append(mapping)
    with open(HISTORY_MAPPINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(history_mappings, f, ensure_ascii=False, indent=4)

# 修改选择模板函数
def select_template():
    global template_df
    history_templates = load_history_templates()
    template_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("历史模板", [os.path.basename(t) for t in history_templates])])
    if template_path:
        try:
            # FIX: 使用 read_excel 替代 ExcelFile.parse()
            template_df = pd.read_excel(template_path, sheet_name=0, dtype=str)  # 读取第一个工作表，保留前导零
            messagebox.showinfo("成功", "模板文件已加载")
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

    # 让用户选择需要转换的文件
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, sheet_name=0, dtype=str)  # 读取第一个工作表，保留前导零
        file_columns = df.columns.tolist()
    except Exception as e:
        messagebox.showerror("错误", f"读取文件 {file_path} 时出错: {str(e)}")
        log_text.insert(tk.END, f"读取文件 {file_path} 出错: {str(e)}\n")
        return

    mapping_window = tk.Toplevel(root)
    mapping_window.title("设置列映射")

    # 常用时间格式选项（使用 yyyy-MM-dd 风格，更直观）
    date_formats = [
        "",  # 不转换
        "yyyyMMdd",             # 20250101
        "yyyy/MM/dd",           # 2025/01/01
        "yyyy-MM-dd",           # 2025-01-01
        "yyyy年MM月dd日",        # 2025年01月01日
        "dd/MM/yyyy",           # 01/01/2025
        "MM/dd/yyyy",           # 01/01/2025
        "yyyy-MM-dd HH:mm:ss",  # 2025-01-01 14:30:00
        "yyyy/MM/dd HH:mm:ss",  # 2025/01/01 14:30:00
        "yyyyMMddHHmmss",       # 20250101143000
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

        # 输入时间格式（可编辑，支持自定义）
        input_date_combo = ttk.Combobox(mapping_window, values=date_formats, width=15)
        input_date_combo.grid(row=i + 1, column=3)
        input_date_combos.append(input_date_combo)

        # 输出时间格式（可编辑，支持自定义）
        output_date_combo = ttk.Combobox(mapping_window, values=date_formats, width=15)
        output_date_combo.grid(row=i + 1, column=4)
        output_date_combos.append(output_date_combo)

    # 修改保存映射函数
    def save_mapping():
        global column_mapping
        # 弹出输入框让用户输入映射名称
        name = simpledialog.askstring("输入映射名称", "请输入当前映射关系的名称：")
        if not name:
            messagebox.showwarning("警告", "未输入映射名称，映射未保存")
            return
        column_mapping = {}
        split_info = {}
        date_format_info = {}  # 时间格式转换配置
        for i, col in enumerate(template_columns):
            selected_col = combo_boxes[i].get()
            split_symbol = entry_boxes[i].get()
            input_fmt = input_date_combos[i].get()
            output_fmt = output_date_combos[i].get()
            if selected_col:
                column_mapping[col] = selected_col
                if split_symbol:
                    split_info[selected_col] = split_symbol
                # 只有同时指定了输入和输出格式才保存
                if input_fmt and output_fmt:
                    date_format_info[col] = {"input": input_fmt, "output": output_fmt}
        column_mapping['split_info'] = split_info
        column_mapping['date_format_info'] = date_format_info
        mapping_window.destroy()
        messagebox.showinfo("成功", "列映射已保存")
        log_text.insert(tk.END, "列映射已成功保存\n")
        # 修改保存历史映射逻辑，使用用户输入的名称作为 key
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
        messagebox.showinfo("成功", f"保存目录已设置为: {save_dir}")
        log_text.insert(tk.END, f"保存目录已设置为: {save_dir}\n")


# 修改转换函数
def convert_excel_files():
    if 'template_df' not in globals() or 'column_mapping' not in globals() or column_mapping is None:
        messagebox.showerror("错误", "请先选择模板并设置列映射")
        return
    if save_dir is None:
        messagebox.showerror("错误", "请先选择保存目录")
        return

    # 选择多个 Excel 文件
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_paths:
        return

    split_info = column_mapping.get('split_info', {})
    date_format_info = column_mapping.get('date_format_info', {})

    # 从 column_mapping 中提取模板列和目标列的映射（排除 split_info 和 date_format_info）
    # column_mapping 格式: {模板列名: 目标列名, ..., 'split_info': {...}, 'date_format_info': {...}}
    template_to_file_mapping = {k: v for k, v in column_mapping.items() if k not in ('split_info', 'date_format_info')}
    
    for file_path in file_paths:
        try:
            # 获取所有工作表名称
            sheet_names = pd.ExcelFile(file_path).sheet_names

            for sheet_name in sheet_names:
                # 读取每个工作表，dtype=str 保留前导零
                df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)

                log_text.insert(tk.END, f"调试信息: 文件 {file_path} 工作表 {sheet_name} - 开始处理\n")
                log_text.insert(tk.END, f"调试信息: column_mapping: {column_mapping}\n")

                # 获取需要提取的目标文件列名（按模板列的顺序）
                file_columns_ordered = [template_to_file_mapping[k] for k in template_to_file_mapping.keys()]
                
                # 检查目标文件中是否包含所需的列
                missing_columns = [col for col in file_columns_ordered if col not in df.columns]
                if missing_columns:
                    log_text.insert(tk.END, f"错误: 文件 {file_path} 工作表 {sheet_name} 缺少列: {missing_columns}\n")
                    continue
                
                # 提取映射的列
                mapped_df = df[file_columns_ordered].copy()
                log_text.insert(tk.END, f"调试信息: 提取的列: {file_columns_ordered}\n")
                
                # 立即重命名列名为模板列名，方便后续处理
                template_columns_ordered = list(template_to_file_mapping.keys())
                mapped_df.columns = template_columns_ordered
                log_text.insert(tk.END, f"调试信息: 列名已重命名为模板列名: {template_columns_ordered}\n")

                # 生成反向列映射: {目标文件列名(小写): 模板列名(小写)}
                reverse_column_mapping = {v.strip().lower(): k.strip().lower() for k, v in template_to_file_mapping.items()}
                
                # 调整 split_info: {模板列名(原始大小写): 分隔符}
                # split_info 中的 key 是目标文件的列名，需要转换为模板列名
                adjusted_split_info = {}
                for file_col, split_symbol in split_info.items():
                    template_col_lower = reverse_column_mapping.get(file_col.strip().lower())
                    if template_col_lower:
                        # 从 template_columns_ordered 中找到原始大小写的列名
                        for orig_col in template_columns_ordered:
                            if orig_col.strip().lower() == template_col_lower:
                                adjusted_split_info[orig_col] = split_symbol
                                break
                    else:
                        log_text.insert(tk.END, f"警告: split_info 中的列 '{file_col}' 未在映射中找到\n")

                log_text.insert(tk.END, f"调试信息: adjusted_split_info: {adjusted_split_info}\n")

                if adjusted_split_info:
                    # 处理分割逻辑
                    new_rows = []
                    for index, row in mapped_df.iterrows():
                        try:
                            # 收集需要分割的列及其分割后的值
                            split_values_dict = {}
                            max_length = 1
                            
                            # 第一步：收集所有需要分割的列和它们的分割结果
                            for col_name, split_symbol in adjusted_split_info.items():
                                if col_name in row.index:
                                    if pd.notna(row[col_name]):
                                        # 按分隔符分割
                                        values = str(row[col_name]).split(split_symbol)
                                        # 去除每个值两端的空白和可能残留的分隔符
                                        values = [v.strip().strip(split_symbol).strip() for v in values]
                                        # 过滤掉空字符串（处理 |xxxx|yyyy 这种开头或结尾有分隔符的情况）
                                        values = [v for v in values if v]
                                        if values:
                                            split_values_dict[col_name] = values
                                            max_length = max(max_length, len(values))
                                            log_text.insert(tk.END, f"调试: 行 {index} 列 '{col_name}' 原始值='{row[col_name]}' 按'{split_symbol}'分割为 {len(values)} 个值: {values}\n")
                                        else:
                                            split_values_dict[col_name] = [str(row[col_name])]
                                    else:
                                        split_values_dict[col_name] = ['']
                                else:
                                    log_text.insert(tk.END, f"警告: 行 {index} 中不存在列 '{col_name}'\n")
                            
                            # 第二步：生成新行
                            if split_values_dict:
                                for i in range(max_length):
                                    new_row = row.copy()
                                    # 为每个需要分割的列赋值
                                    for col_name, values in split_values_dict.items():
                                        # 如果该列的值数量少于max_length，则使用空字符串填充
                                        if i < len(values):
                                            new_row[col_name] = values[i]
                                        else:
                                            new_row[col_name] = ''
                                    new_rows.append(new_row)
                                    
                                log_text.insert(tk.END, f"信息: 行 {index} 拆分为 {max_length} 行\n")
                            else:
                                # 如果没有需要分割的数据或max_length为1，保留原行
                                new_rows.append(row)

                        except Exception as row_error:
                            log_text.insert(tk.END, f"错误: 处理行 {index} 时出错: {str(row_error)}\n")
                            import traceback
                            log_text.insert(tk.END, f"详细错误: {traceback.format_exc()}\n")
                            # 出错时保留原行
                            new_rows.append(row)
                            continue

                    mapped_df = pd.DataFrame(new_rows)
                    log_text.insert(tk.END, f"信息: 拆分完成，共 {len(new_rows)} 行\n")
                else:
                    log_text.insert(tk.END, f"信息: 无需拆分\n")

                # 时间格式转换
                if date_format_info:
                    from datetime import datetime

                    def convert_date_format(fmt):
                        """将 yyyy-MM-dd 风格转换为 Python strftime 格式"""
                        if not fmt:
                            return fmt
                        # 顺序很重要：先替换长的，再替换短的
                        replacements = [
                            ('yyyy', '%Y'),
                            ('yy', '%y'),
                            ('MM', '%m'),
                            ('dd', '%d'),
                            ('HH', '%H'),
                            ('mm', '%M'),
                            ('ss', '%S'),
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
                                        # 格式不匹配时保留原值
                                        return str(val)
                                mapped_df[col] = mapped_df[col].apply(convert_date)
                                log_text.insert(tk.END, f"信息: 列 '{col}' 时间格式已从 '{fmt_config.get('input', '')}' 转换为 '{fmt_config.get('output', '')}'\n")

                # 确保最终的 DataFrame 包含模板的所有列
                all_template_columns = list(template_df.columns)
                
                # 为缺失的列添加空值
                for col in all_template_columns:
                    if col not in mapped_df.columns:
                        mapped_df[col] = ''
                        log_text.insert(tk.END, f"信息: 添加缺失的模板列 '{col}'\n")
                
                # 按照模板列的顺序重新排列列
                mapped_df = mapped_df[all_template_columns]
                log_text.insert(tk.END, f"信息: 最终列顺序与模板一致: {all_template_columns}\n")
                
                # 生成保存的文件名
                file_name = os.path.splitext(os.path.basename(file_path))[0]
                if len(sheet_names) > 1:
                    file_name += f'_{sheet_name}'
                save_path = os.path.join(save_dir, f'{file_name}.xlsx')

                # 将所有列转换为字符串类型，确保以文本格式保存
                for col in mapped_df.columns:
                    mapped_df[col] = mapped_df[col].astype(str)
                    # 将 'nan' 字符串替换为空字符串
                    mapped_df[col] = mapped_df[col].replace('nan', '')

                # 将数据保存为 Excel 文件，添加文件锁检测
                try:
                    # 使用 openpyxl 引擎，所有数据作为文本写入
                    from openpyxl import load_workbook
                    from openpyxl.styles import numbers
                    
                    # 先保存为临时 Excel 文件
                    mapped_df.to_excel(save_path, index=False, sheet_name='Sheet1', engine='openpyxl')
                    
                    # 重新打开文件，将所有单元格格式设置为文本
                    wb = load_workbook(save_path)
                    ws = wb.active
                    
                    # 设置所有数据单元格为文本格式（从第2行开始，第1行是表头）
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                        for cell in row:
                            cell.number_format = numbers.FORMAT_TEXT
                    
                    # 保存文件
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
            # 添加详细的错误日志
            log_text.insert(tk.END, f"错误: 处理文件 {file_path} 时出错，详细错误信息: {str(e)}\n")
            messagebox.showerror("错误", f"处理文件 {file_path} 时出错: {str(e)}")

    messagebox.showinfo("完成", "所有文件转换完成！")


# 显示历史模板
def show_history_templates():
    history_templates = load_history_templates()
    if not history_templates:
        messagebox.showinfo("历史模板", "暂无历史模板记录")
    else:
        template_window = tk.Toplevel(root)
        template_window.title("历史模板")
        for template_path in history_templates:
            tk.Button(template_window, text=template_path, command=lambda path=template_path: use_history_template(path)).pack(pady=5)

# 使用历史模板
def use_history_template(template_path):
    global template_df
    try:
        # FIX: 使用 read_excel 替代 ExcelFile.parse()
        template_df = pd.read_excel(template_path, sheet_name=0, dtype=str)  # 保留前导零
        messagebox.showinfo("成功", "历史模板文件已加载")
        log_text.insert(tk.END, f"历史模板文件 {template_path} 已成功加载\n")
        save_history_template(template_path)
    except Exception as e:
        messagebox.showerror("错误", f"加载历史模板文件时出错: {str(e)}")
        log_text.insert(tk.END, f"加载历史模板文件 {template_path} 出错: {str(e)}\n")

# 显示历史映射
def show_history_mappings():
    history_mappings = load_history_mappings()
    if not history_mappings:
        messagebox.showinfo("历史映射", "暂无历史映射记录")
    else:
        mapping_window = tk.Toplevel(root)
        mapping_window.title("历史映射")
        for name in history_mappings:
            tk.Button(mapping_window, text=name, command=lambda n=name: use_history_mapping(history_mappings[n])).pack(pady=5)

# 使用历史映射
def use_history_mapping(mapping):
    global column_mapping
    column_mapping = mapping
    messagebox.showinfo("成功", "历史映射已应用")
    log_text.insert(tk.END, "历史映射已成功应用\n")

# 创建主窗口
root = tk.Tk()
root.title("Excel 文件转换器")

# 选择模板按钮
template_button = tk.Button(root, text="选择模板文件", command=select_template)
template_button.pack(pady=10)

# 设置列映射按钮
mapping_button = tk.Button(root, text="设置列映射", command=set_column_mapping)
mapping_button.pack(pady=10)

# 选择保存目录按钮
select_dir_button = tk.Button(root, text="选择生成文件路径", command=select_save_directory)
select_dir_button.pack(pady=10)

# 创建转换按钮
convert_button = tk.Button(root, text="转换 Excel 文件", command=convert_excel_files)
convert_button.pack(pady=10)

# 显示历史模板按钮
show_templates_button = tk.Button(root, text="显示历史模板", command=show_history_templates)
show_templates_button.pack(pady=10)

# 显示历史映射按钮
show_mappings_button = tk.Button(root, text="显示历史映射", command=show_history_mappings)
show_mappings_button.pack(pady=10)

# 运行日志展示区域
log_text = scrolledtext.ScrolledText(root, width=50, height=10)
log_text.pack(pady=10)

# 启动主循环
root.mainloop()