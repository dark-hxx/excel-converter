"""银行流水转换：核心逻辑 + 单文件窗口 + 批量混合窗口。"""
import json
import os
import re
import sys
import tkinter as tk
from datetime import datetime
from decimal import Decimal
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

from utils import center_window, open_folder


# ---------------- 路径与常量 ----------------

if getattr(sys, 'frozen', False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.getcwd()

# 规则文件：打包后在 sys._MEIPASS，开发期在工作目录
BANK_RULES_FILE = os.path.join(
    getattr(sys, '_MEIPASS', _BASE_DIR), 'bank_rules.json'
)

# 用户上次选择（银行 + 保存目录）
LAST_CHOICE_FILE = os.path.join(_BASE_DIR, 'last_choice.json')

# 统一模板表头（27 列，严格匹配后端 ImportBankDetailDTO.getHeadList()）
BANK_TEMPLATE_HEADERS = [
    '交易后余额', '账户名称', '账户性质', '银行账号', '关联客户号',
    '银行流水号', '交易日期', '银行类型', '对账码', '币种',
    '摘要', '业务参考号', '借方金额', '贷方金额', '客商名称',
    '客商编号', '开户行名称', '对方账号', '对方户名', '对方开户行',
    '款项性质代码', '用途', '备注', '交易流水号', '单位名称',
    '起息日', '借/贷金额',
]

# 文件名 → 银行类型 关键字映射（按 key 长度倒序匹配，避免「中行」吞掉「中信」）
BANK_FILENAME_KEYWORDS = {
    '招商银行': '招商银行', '招商': '招商银行', 'CMB': '招商银行',
    '工商银行': '工商银行', '工行': '工商银行', 'ICBC': '工商银行',
    '农业银行': '农业银行', '农行': '农业银行', 'ABC': '农业银行',
    '中国银行': '中国银行', '中行': '中国银行', 'BOC': '中国银行',
    '交通银行': '交通银行', '交行': '交通银行', 'BCM': '交通银行',
    '中信银行': '中信银行', '中信': '中信银行', 'CITIC': '中信银行',
    '民生银行': '民生银行', '民生': '民生银行', 'CMBC': '民生银行',
    '北京银行': '北京银行', '北行': '北京银行', 'BJB': '北京银行',
    '上海银行': '上海银行', '上行': '上海银行', 'BOSC': '上海银行',
    '南京银行': '南京银行', '南京': '南京银行', 'NJB': '南京银行',
    '宁波银行': '宁波银行', '宁波': '宁波银行', 'NBB': '宁波银行',
    '农商银行': '农商银行', '农商': '农商银行', 'NSB': '农商银行',
    '江西银行': '江西银行', '江西': '江西银行', 'JXB': '江西银行',
}


# ---------------- 配置加载 ----------------

def load_bank_rules():
    """加载银行规则配置。失败时弹错并返回空字典。"""
    try:
        with open(BANK_RULES_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return {k: v for k, v in data.items() if not k.startswith('_')}
    except FileNotFoundError:
        messagebox.showerror('错误', f'未找到银行规则文件: {BANK_RULES_FILE}')
        return {}
    except Exception as e:
        messagebox.showerror('错误', f'加载银行规则失败: {e}')
        return {}


def load_last_choice():
    try:
        with open(LAST_CHOICE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_last_choice(bank, save_dir_path):
    try:
        with open(LAST_CHOICE_FILE, 'w', encoding='utf-8') as f:
            json.dump({'bank': bank, 'save_dir': save_dir_path},
                      f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def guess_bank_by_filename(filename, available_banks):
    """根据文件名启发式猜银行；猜不到返回 None。available_banks 限定候选范围。"""
    name = os.path.basename(filename)
    # 按 key 长度倒序，长关键字优先（避免「中行」匹配到「中信银行」）
    for kw in sorted(BANK_FILENAME_KEYWORDS.keys(), key=len, reverse=True):
        if kw in name:
            bank = BANK_FILENAME_KEYWORDS[kw]
            if bank in available_banks:
                return bank
    return None


# ---------------- 转换核心 ----------------

def _bank_log(log_widget, msg):
    """日志输出（带自动滚动与刷新）"""
    if log_widget is None:
        return
    log_widget.insert(tk.END, msg + '\n')
    log_widget.see(tk.END)
    log_widget.update_idletasks()


def parse_amount(val):
    """解析金额：去千分位、去货币符号；空/横线/异常返回 None"""
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in ('-', '--', 'nan', 'None'):
        return None
    s = s.replace(',', '').replace(' ', '').replace('¥', '').replace('￥', '')
    try:
        return Decimal(s)
    except Exception:
        return None


def _amount_to_str(d):
    if d is None:
        return ''
    s = format(d, 'f')
    if '.' in s:
        s = s.rstrip('0').rstrip('.')
    return s or '0'


def extract_account_info(file_path, rule):
    """从顶部行/单元格提取账户信息（账号、户名等）。返回 dict。"""
    extract_cfg = rule.get('account_extract')
    if not extract_cfg:
        return {}

    from openpyxl import load_workbook
    try:
        wb = load_workbook(file_path, data_only=True, read_only=True)
        ws = wb.worksheets[0]  # 与 pd.read_excel(sheet_name=0) 保持一致
        result = {}
        for tpl_field, spec in extract_cfg.items():
            cell_addr = spec.get('cell')
            if not cell_addr:
                continue
            raw = ws[cell_addr].value
            if raw is None:
                continue
            text = str(raw)
            if spec.get('strip', False):
                text = text.strip().strip('\t').strip()
            if spec.get('mode', 'cell') == 'regex':
                m = re.search(spec.get('pattern', ''), text)
                if m and m.groups():
                    text = m.group(1).strip()
            result[tpl_field] = text
        wb.close()
        return result
    except Exception:
        return {}


def _get_cell(row, col_name):
    """从 Series/dict 中安全取值；列不存在返回 None"""
    if col_name is None:
        return None
    try:
        if (col_name in row.index) if hasattr(row, 'index') else (col_name in row):
            v = row[col_name]
            if pd.isna(v):
                return None
            return v
    except Exception:
        pass
    return None


def build_date_value(row, spec):
    """根据 spec 构建日期字符串。spec 可以是 str (列名) 或 dict。"""
    if isinstance(spec, str):
        val = _get_cell(row, spec)
        return '' if val is None else str(val).strip()

    source = spec.get('source')
    in_fmt = spec.get('in_fmt')
    out_fmt = spec.get('out_fmt')
    join_str = spec.get('join', ' ')

    if isinstance(source, list):
        parts = []
        for col in source:
            v = _get_cell(row, col)
            if v is None:
                parts.append('')
                continue
            s = str(v).strip()
            # 民生银行：A 列含 "2026-04-06 00:00:00"，只取日期部分
            if spec.get('date_only_first') and len(parts) == 0 and ' ' in s:
                s = s.split(' ')[0]
            parts.append(s)
        raw = join_str.join([p for p in parts if p])
    else:
        v = _get_cell(row, source)
        raw = '' if v is None else str(v).strip()

    if not raw:
        return ''

    # 农商银行：时间 HH:mm 缺秒，补 :00
    if spec.get('pad_seconds') and raw.count(':') == 1:
        raw = raw + ':00'

    if not in_fmt or not out_fmt:
        return raw

    try:
        return datetime.strptime(raw, in_fmt).strftime(out_fmt)
    except Exception:
        return raw  # 解析失败保留原值，让后端校验报错


def resolve_field(row, spec):
    """根据 spec 解析普通字段值（非日期、非金额）"""
    if spec is None:
        return ''
    if isinstance(spec, str):
        v = _get_cell(row, spec)
        return '' if v is None else str(v).strip()
    src = spec.get('source')
    v = _get_cell(row, src)
    if v is None:
        return ''
    s = str(v).strip()
    if spec.get('strip_spaces'):
        s = s.replace(' ', '')
    return s


def normalize_debit_credit(row, rule):
    """借贷归一，返回 (借方金额字符串, 贷方金额字符串)"""
    mode = rule.get('debit_credit_mode', 'two_columns')
    cm = rule.get('column_mapping', {})

    if mode == 'two_columns':
        jie_spec = cm.get('借方金额')
        dai_spec = cm.get('贷方金额')
        jie_col = jie_spec if isinstance(jie_spec, str) else (jie_spec or {}).get('source')
        dai_col = dai_spec if isinstance(dai_spec, str) else (dai_spec or {}).get('source')
        jie = parse_amount(_get_cell(row, jie_col))
        dai = parse_amount(_get_cell(row, dai_col))
        return _amount_to_str(jie), _amount_to_str(dai)

    if mode == 'signed_amount':
        amt = parse_amount(_get_cell(row, rule.get('amount_column')))
        if amt is None or amt == 0:
            return '', ''
        if amt < 0:
            return _amount_to_str(-amt), ''
        return '', _amount_to_str(amt)

    if mode == 'marker_column':
        mc = rule.get('marker_column', {})
        marker = _get_cell(row, mc.get('col'))
        marker_str = '' if marker is None else str(marker).strip()
        amt = parse_amount(_get_cell(row, mc.get('amount_col')))
        if amt is None:
            return '', ''
        if marker_str == mc.get('jie_value'):
            return _amount_to_str(amt), ''
        if marker_str == mc.get('dai_value'):
            return '', _amount_to_str(amt)
        return '', ''

    return '', ''


def _transform_one_row(row, rule, account_info, bank_type_code):
    """把源文件一行转成目标模板的 dict（27 列）。失败字段留空。"""
    cm = rule.get('column_mapping', {})
    fixed = rule.get('fixed_values', {})

    out = {h: '' for h in BANK_TEMPLATE_HEADERS}

    for tpl_field, spec in cm.items():
        if tpl_field in ('交易日期', '起息日'):
            out[tpl_field] = build_date_value(row, spec)
        elif tpl_field in ('借方金额', '贷方金额'):
            pass  # 由 normalize 统一处理
        elif tpl_field == '交易后余额':
            bal_col = spec if isinstance(spec, str) else spec.get('source')
            out[tpl_field] = _amount_to_str(parse_amount(_get_cell(row, bal_col)))
        else:
            out[tpl_field] = resolve_field(row, spec)

    jie, dai = normalize_debit_credit(row, rule)
    out['借方金额'] = jie
    out['贷方金额'] = dai

    for k, v in account_info.items():
        if not out.get(k):
            out[k] = v
    for k, v in fixed.items():
        if not out.get(k):
            out[k] = str(v)
    if not out.get('银行类型'):
        out['银行类型'] = bank_type_code

    return out


def convert_bank_rows(file_path, rule, log_widget=None):
    """读取并转换一个文件，返回 (rows_list, skipped_count, error_msg)。

    error_msg 非空表示失败；rows_list 可能为空（表示无有效行）。
    """
    try:
        header_row = rule.get('header_row', 1)
        df = pd.read_excel(file_path, sheet_name=0,
                           header=header_row - 1, dtype=str)
    except Exception as e:
        return [], 0, f'读取失败: {e}'

    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]
    _bank_log(log_widget, f'  共 {len(df)} 行待处理')

    account_info = extract_account_info(file_path, rule)
    if account_info:
        _bank_log(log_widget, f'  顶部账户信息: {account_info}')

    bank_type_code = rule.get('bank_type_code', '')
    rows_out = []
    skipped = 0

    for idx, row in df.iterrows():
        out = _transform_one_row(row, rule, account_info, bank_type_code)
        if not out['交易后余额'] or not out['交易日期']:
            skipped += 1
            _bank_log(log_widget,
                      f'  跳过第 {idx + header_row + 1} 行：余额或交易日期为空')
            continue
        rows_out.append(out)

    return rows_out, skipped, ''


def _write_bank_xlsx(rows, save_path):
    """把 rows 写入指定 xlsx，所有单元格按文本格式。失败抛异常。"""
    from openpyxl import load_workbook
    from openpyxl.styles import numbers

    out_df = pd.DataFrame(rows, columns=BANK_TEMPLATE_HEADERS)
    out_df.to_excel(save_path, index=False, sheet_name='Sheet1', engine='openpyxl')
    wb = load_workbook(save_path)
    ws = wb.active
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in r:
            cell.number_format = numbers.FORMAT_TEXT
    wb.save(save_path)
    wb.close()


def convert_bank_file(file_path, bank_name, rule, save_dir_path, log_widget):
    """单文件转换（写出独立文件）。成功返回 True。"""
    _bank_log(log_widget, f'开始处理 [{bank_name}] {file_path}')

    rows, skipped, err = convert_bank_rows(file_path, rule, log_widget)
    if err:
        _bank_log(log_widget, f'  {err}')
        return False
    if not rows:
        _bank_log(log_widget, '  无有效数据行，未生成文件')
        return False

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    save_path = os.path.join(save_dir_path, f'{file_name}_{bank_name}_统一格式.xlsx')

    try:
        _write_bank_xlsx(rows, save_path)
    except PermissionError:
        _bank_log(log_widget, f'  保存失败：{save_path} 被占用，请关闭后重试')
        return False
    except Exception as e:
        _bank_log(log_widget, f'  保存失败: {e}')
        return False

    _bank_log(log_widget,
              f'  完成：输出 {len(rows)} 行，跳过 {skipped} 行 → {save_path}')
    return True


def preview_bank_file(file_path, bank_name, rule, win_parent, log_widget):
    """读取前 5 行做转换预演，弹窗用 Treeview 展示关键字段。不写文件。"""
    try:
        header_row = rule.get('header_row', 1)
        df = pd.read_excel(file_path, sheet_name=0,
                           header=header_row - 1, dtype=str, nrows=5)
    except Exception as e:
        _bank_log(log_widget, f'预览失败: {e}')
        return

    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]

    account_info = extract_account_info(file_path, rule)
    bank_type_code = rule.get('bank_type_code', '')

    preview_rows = [
        _transform_one_row(row, rule, account_info, bank_type_code)
        for _, row in df.iterrows()
    ]
    if not preview_rows:
        _bank_log(log_widget, '无数据可预览')
        return

    preview_cols = ['交易日期', '银行账号', '银行类型', '币种', '借方金额',
                    '贷方金额', '交易后余额', '对方账号', '对方户名', '摘要']
    col_widths = {
        '交易日期': 150, '银行账号': 180, '银行类型': 70, '币种': 50,
        '借方金额': 110, '贷方金额': 110, '交易后余额': 130,
        '对方账号': 180, '对方户名': 200, '摘要': 180,
    }

    pw = tk.Toplevel(win_parent)
    pw.title(f'预览：{bank_name} - 前 {len(preview_rows)} 行')
    center_window(pw, 1280, 360)
    pw.transient(win_parent)
    pw.grab_set()

    container = tk.Frame(pw)
    container.pack(fill='both', expand=True, padx=10, pady=10)

    style = ttk.Style()
    style.configure('Preview.Treeview', rowheight=28)

    tree = ttk.Treeview(container, columns=preview_cols, show='headings',
                        height=len(preview_rows), style='Preview.Treeview')
    for col in preview_cols:
        tree.heading(col, text=col)
        tree.column(col, width=col_widths.get(col, 120), anchor='w', stretch=False)
    for r in preview_rows:
        tree.insert('', 'end', values=[r.get(c, '') for c in preview_cols])

    vsb = ttk.Scrollbar(container, orient='vertical', command=tree.yview)
    hsb = ttk.Scrollbar(container, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    container.rowconfigure(0, weight=1)
    container.columnconfigure(0, weight=1)

    tk.Button(pw, text='关闭', width=10, command=pw.destroy).pack(pady=(0, 10))


# ---------------- 单文件转换窗口 ----------------

def open_bank_converter_window(parent_root):
    """打开「银行流水转换」单文件窗口"""
    rules = load_bank_rules()
    if not rules:
        return

    last = load_last_choice()
    bank_names = list(rules.keys())
    last_bank = last.get('bank') if last.get('bank') in bank_names else bank_names[0]
    last_dir = last.get('save_dir') or ''

    win = tk.Toplevel(parent_root)
    win.title('银行流水转换')
    center_window(win, 680, 560)
    win.transient(parent_root)
    win.grab_set()
    win.focus_set()

    # 银行类型
    tk.Label(win, text='银行类型：').grid(row=0, column=0, sticky='w', padx=10, pady=8)
    bank_var = tk.StringVar()
    bank_combo = ttk.Combobox(win, textvariable=bank_var,
                              values=bank_names, state='readonly', width=30)
    bank_combo.grid(row=0, column=1, sticky='w', padx=10, pady=8)
    bank_combo.set(last_bank)

    # 文件选择
    file_var = tk.StringVar()
    tk.Label(win, text='源文件：').grid(row=1, column=0, sticky='w', padx=10, pady=8)
    tk.Label(win, textvariable=file_var, fg='blue',
             wraplength=350, justify='left', anchor='w'
             ).grid(row=1, column=1, sticky='w', padx=10, pady=8)

    def pick_file():
        p = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx;*.xls')])
        if p:
            file_var.set(p)

    tk.Button(win, text='选择文件', command=pick_file).grid(row=1, column=2, padx=5)

    # 保存目录
    dir_var = tk.StringVar(value=last_dir)
    tk.Label(win, text='保存目录：').grid(row=2, column=0, sticky='w', padx=10, pady=8)
    tk.Label(win, textvariable=dir_var, fg='blue',
             wraplength=350, justify='left', anchor='w'
             ).grid(row=2, column=1, sticky='w', padx=10, pady=8)

    def pick_dir():
        d = filedialog.askdirectory()
        if d:
            dir_var.set(d)

    tk.Button(win, text='选择目录', command=pick_dir).grid(row=2, column=2, padx=5)

    # 日志区
    log_widget = scrolledtext.ScrolledText(win, width=70, height=18)
    log_header = tk.Frame(win)
    log_header.grid(row=4, column=0, columnspan=3, sticky='ew', padx=10, pady=(8, 0))
    tk.Label(log_header, text='日志：').pack(side='left')
    tk.Button(log_header, text='清空', width=6,
              command=lambda: log_widget.delete('1.0', tk.END)).pack(side='right')
    log_widget.grid(row=5, column=0, columnspan=3, padx=10, pady=(0, 10))

    # 操作区
    btn_frame = tk.Frame(win)
    btn_frame.grid(row=3, column=0, columnspan=3, pady=10)
    buttons = {}

    def do_preview():
        bank_name = bank_var.get()
        file_path = file_var.get()
        if not bank_name:
            messagebox.showwarning('提示', '请选择银行类型')
            return
        if not file_path:
            messagebox.showwarning('提示', '请选择源文件')
            return
        rule = rules.get(bank_name)
        if rule:
            preview_bank_file(file_path, bank_name, rule, win, log_widget)

    def do_convert():
        bank_name = bank_var.get()
        file_path = file_var.get()
        save_path = dir_var.get()

        if not bank_name:
            messagebox.showwarning('提示', '请选择银行类型')
            return
        if not file_path:
            messagebox.showwarning('提示', '请选择源文件')
            return
        if not save_path:
            messagebox.showwarning('提示', '请选择保存目录')
            return

        rule = rules.get(bank_name)
        if not rule:
            messagebox.showerror('错误', f'未找到银行规则：{bank_name}')
            return

        log_widget.delete('1.0', tk.END)
        buttons['preview'].config(state='disabled')
        buttons['convert'].config(state='disabled', text='处理中...')
        win.update_idletasks()
        try:
            ok = convert_bank_file(file_path, bank_name, rule, save_path, log_widget)
        finally:
            buttons['preview'].config(state='normal')
            buttons['convert'].config(state='normal', text='开始转换')

        if ok:
            save_last_choice(bank_name, save_path)
            if messagebox.askyesno('完成', '转换完成，是否打开输出目录？'):
                open_folder(save_path)
        else:
            messagebox.showwarning('提示', '转换未完成，请查看日志')

    buttons['preview'] = tk.Button(btn_frame, text='预览前 5 行', command=do_preview, width=15)
    buttons['preview'].pack(side='left', padx=5)
    buttons['convert'] = tk.Button(btn_frame, text='开始转换', command=do_convert,
                                   bg='#4CAF50', fg='white', width=15)
    buttons['convert'].pack(side='left', padx=5)


# ---------------- 批量混合窗口 ----------------

def open_batch_converter_window(parent_root):
    """打开「批量混合转换」窗口：多文件 + 每文件单独选银行 + 合并到一个 xlsx"""
    rules = load_bank_rules()
    if not rules:
        return

    bank_names = list(rules.keys())
    last = load_last_choice()
    last_dir = last.get('save_dir') or ''

    win = tk.Toplevel(parent_root)
    win.title('批量混合转换')
    center_window(win, 920, 680)
    win.transient(parent_root)
    win.grab_set()
    win.focus_set()

    # 内部状态：iid → 完整文件路径
    path_by_iid = {}

    # ---- 文件操作按钮 ----
    file_btn_row = tk.Frame(win)
    file_btn_row.pack(fill='x', padx=10, pady=(10, 4))

    def add_files():
        paths = filedialog.askopenfilenames(filetypes=[('Excel files', '*.xlsx;*.xls')])
        for p in paths:
            bank = guess_bank_by_filename(p, bank_names) or bank_names[0]
            iid = tree.insert('', 'end', values=(
                len(path_by_iid) + 1, os.path.basename(p), bank, '待转换'))
            path_by_iid[iid] = p
        _renumber()

    def remove_selected():
        for iid in tree.selection():
            path_by_iid.pop(iid, None)
            tree.delete(iid)
        _renumber()

    def clear_all():
        for iid in list(path_by_iid):
            tree.delete(iid)
        path_by_iid.clear()

    def _renumber():
        for i, iid in enumerate(tree.get_children(), 1):
            vals = list(tree.item(iid, 'values'))
            vals[0] = i
            tree.item(iid, values=vals)

    tk.Button(file_btn_row, text='+ 添加文件', command=add_files,
              bg='#4CAF50', fg='white', width=12).pack(side='left', padx=2)
    tk.Button(file_btn_row, text='- 移除选中', command=remove_selected, width=12).pack(side='left', padx=2)
    tk.Button(file_btn_row, text='清空', command=clear_all, width=10).pack(side='left', padx=2)
    tk.Label(file_btn_row, text='（多选支持 Ctrl/Shift 点击）', fg='#888').pack(side='left', padx=8)

    # ---- 文件列表 Treeview ----
    list_container = tk.Frame(win)
    list_container.pack(fill='both', expand=True, padx=10, pady=4)

    list_cols = ('#', '文件名', '银行类型', '状态')
    col_widths = {'#': 40, '文件名': 380, '银行类型': 140, '状态': 140}

    style = ttk.Style()
    style.configure('Batch.Treeview', rowheight=26)

    tree = ttk.Treeview(list_container, columns=list_cols, show='headings',
                        style='Batch.Treeview', height=12)
    for c in list_cols:
        tree.heading(c, text=c)
        tree.column(c, width=col_widths.get(c, 100),
                    anchor='center' if c in ('#', '状态') else 'w', stretch=(c == '文件名'))

    vsb = ttk.Scrollbar(list_container, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    list_container.rowconfigure(0, weight=1)
    list_container.columnconfigure(0, weight=1)

    # ---- 修改选中行的银行 ----
    edit_row = tk.Frame(win)
    edit_row.pack(fill='x', padx=10, pady=4)
    tk.Label(edit_row, text='选中行改为：').pack(side='left')
    edit_var = tk.StringVar(value=bank_names[0])
    ttk.Combobox(edit_row, textvariable=edit_var, values=bank_names,
                 state='readonly', width=20).pack(side='left', padx=4)

    def apply_bank():
        sels = tree.selection()
        if not sels:
            messagebox.showinfo('提示', '请先在列表中选中行（可多选）')
            return
        bank = edit_var.get()
        for iid in sels:
            vals = list(tree.item(iid, 'values'))
            vals[2] = bank
            vals[3] = '待转换'
            tree.item(iid, values=vals)

    tk.Button(edit_row, text='应用', command=apply_bank, width=8).pack(side='left', padx=4)

    # ---- 输出设置 ----
    out_row = tk.Frame(win)
    out_row.pack(fill='x', padx=10, pady=(8, 4))

    tk.Label(out_row, text='输出目录：').grid(row=0, column=0, sticky='w', pady=2)
    out_dir_var = tk.StringVar(value=last_dir)
    tk.Label(out_row, textvariable=out_dir_var, fg='blue',
             wraplength=600, justify='left', anchor='w'
             ).grid(row=0, column=1, sticky='w', padx=6)

    def pick_out_dir():
        d = filedialog.askdirectory()
        if d:
            out_dir_var.set(d)

    tk.Button(out_row, text='浏览', command=pick_out_dir, width=8
              ).grid(row=0, column=2, padx=4)

    tk.Label(out_row, text='输出文件名：').grid(row=1, column=0, sticky='w', pady=2)
    out_name_var = tk.StringVar(
        value=f'合并流水_{datetime.now().strftime("%Y%m%d")}.xlsx')
    tk.Entry(out_row, textvariable=out_name_var, width=50
             ).grid(row=1, column=1, sticky='w', padx=6, pady=2)
    out_row.columnconfigure(1, weight=1)

    # ---- 日志区 ----
    log_widget = scrolledtext.ScrolledText(win, height=10)
    log_header = tk.Frame(win)
    log_header.pack(fill='x', padx=10, pady=(8, 0))
    tk.Label(log_header, text='日志：').pack(side='left')
    tk.Button(log_header, text='清空', width=6,
              command=lambda: log_widget.delete('1.0', tk.END)).pack(side='right')
    log_widget.pack(fill='both', expand=False, padx=10, pady=(0, 6))

    # ---- 开始转换 ----
    action_row = tk.Frame(win)
    action_row.pack(pady=8)

    def do_batch_convert():
        if not path_by_iid:
            messagebox.showwarning('提示', '请先添加文件')
            return
        save_dir_path = out_dir_var.get().strip()
        if not save_dir_path:
            messagebox.showwarning('提示', '请选择输出目录')
            return
        out_name = out_name_var.get().strip()
        if not out_name:
            messagebox.showwarning('提示', '请填写输出文件名')
            return
        if not out_name.lower().endswith('.xlsx'):
            out_name += '.xlsx'

        log_widget.delete('1.0', tk.END)
        convert_btn.config(state='disabled', text='处理中...')
        win.update_idletasks()

        merged = []
        try:
            for iid in tree.get_children():
                vals = list(tree.item(iid, 'values'))
                idx_no, fname, bank, _ = vals
                fpath = path_by_iid[iid]

                tree.item(iid, values=(idx_no, fname, bank, '处理中...'))
                win.update_idletasks()
                _bank_log(log_widget, f'\n[{idx_no}] [{bank}] {fname}')

                rule = rules.get(bank)
                if not rule:
                    tree.item(iid, values=(idx_no, fname, bank, '失败'))
                    _bank_log(log_widget, f'  未找到规则：{bank}，整批终止')
                    messagebox.showerror('整批终止', f'未找到银行规则：{bank}')
                    return

                rows, skipped, err = convert_bank_rows(fpath, rule, log_widget)
                if err:
                    tree.item(iid, values=(idx_no, fname, bank, '失败'))
                    _bank_log(log_widget, f'  {err}，整批终止')
                    messagebox.showerror('整批终止',
                                         f'文件 [{fname}] 转换失败：{err}\n'
                                         f'已合并行数：{len(merged)}（未写入文件）')
                    return
                if not rows:
                    tree.item(iid, values=(idx_no, fname, bank, '无有效行'))
                    _bank_log(log_widget, '  无有效行，跳过但继续后续文件')
                    continue

                merged.extend(rows)
                tree.item(iid, values=(idx_no, fname, bank, f'成功 ({len(rows)} 行)'))
                _bank_log(log_widget, f'  汇入 {len(rows)} 行（跳过 {skipped}）')

            if not merged:
                _bank_log(log_widget, '\n所有文件均无有效数据，未生成合并文件')
                messagebox.showwarning('提示', '所有文件均无有效数据，未生成合并文件')
                return

            save_path = os.path.join(save_dir_path, out_name)
            try:
                _write_bank_xlsx(merged, save_path)
            except PermissionError:
                _bank_log(log_widget, f'\n保存失败：{save_path} 被占用，请关闭后重试')
                messagebox.showerror('保存失败', f'{save_path} 被占用，请关闭后重试')
                return
            except Exception as e:
                _bank_log(log_widget, f'\n保存失败：{e}')
                messagebox.showerror('保存失败', str(e))
                return

            _bank_log(log_widget,
                      f'\n========== 全部完成：合并 {len(merged)} 行 → {save_path} ==========')
            save_last_choice(last.get('bank') or bank_names[0], save_dir_path)
            if messagebox.askyesno('完成',
                                   f'共合并 {len(merged)} 行到 {out_name}\n是否打开输出目录？'):
                open_folder(save_dir_path)
        finally:
            convert_btn.config(state='normal', text='开始转换')

    convert_btn = tk.Button(action_row, text='开始转换', command=do_batch_convert,
                            bg='#4CAF50', fg='white', width=18, height=2)
    convert_btn.pack()
