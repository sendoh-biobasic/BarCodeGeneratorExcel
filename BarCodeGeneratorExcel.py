import pandas as pd
import math
import tkinter as tk
from tkinter import filedialog, messagebox
import os


# ========== 日期解析 (保持不变) ==========
def parse_date_flexible(x):
    if pd.isna(x):
        return None
    try:
        if hasattr(x, "date") and not isinstance(x, str):
            return pd.to_datetime(x).date()
    except:
        pass
    try:
        if isinstance(x, (int, float)) and not isinstance(x, bool) and not math.isnan(x):
            try:
                return pd.to_datetime(x, origin='1899-12-30', unit='D').date()
            except:
                pass
    except:
        pass
    s = str(x).strip()
    for cand in [s, s.replace('.', '-'), s.replace('/', '-'),
                 s.replace('年', '-').replace('月', '-').replace('日', '')]:
        try:
            ts = pd.to_datetime(cand, errors='coerce')
            if not pd.isna(ts):
                return ts.date()
        except:
            continue
    return None


# ========== 主逻辑 (已更新列索引 + 支持 xlsm) ==========
def run_extraction(main_path, db_path, out_path, input_date):
    try:
        df_main = pd.read_excel(main_path, sheet_name="3200 MIGO", dtype=object, engine='openpyxl')
        df_main['ReceivingDate_norm'] = df_main.iloc[:, 0].apply(parse_date_flexible)

        df_db = pd.read_excel(db_path, sheet_name="product-db", dtype=object, engine='openpyxl')

        db_map = {}
        if df_db.shape[1] >= 5:
            for _, r in df_db.iterrows():
                key = "" if pd.isna(r.iloc[4]) else str(r.iloc[4]).strip().lower()
                if key:
                    db_map[key] = (
                        "" if pd.isna(r.iloc[0]) else str(r.iloc[0]),
                        "" if pd.isna(r.iloc[3]) else str(r.iloc[3]),
                    )

        target_date = parse_date_flexible(input_date)
        if not target_date:
            messagebox.showerror("错误", "无法识别输入日期，请确认格式。")
            return

        matched = df_main[df_main['ReceivingDate_norm'] == target_date].reset_index(drop=True)
        if matched.empty:
            messagebox.showwarning("提示", "未找到匹配的 Receiving Date。")
            return

        records = []
        for _, row in matched.iterrows():
            date_val = row.iloc[0] if len(row) > 0 else ""
            item_val = row.iloc[3] if len(row) > 3 else ""
            sap_val = row.iloc[4] if len(row) > 4 else ""
            desc_val = row.iloc[6] if len(row) > 6 else ""  # 新索引

            g_val = "" if len(row) <= 7 or pd.isna(row.iloc[7]) else str(row.iloc[7]).strip()
            h_val = "" if len(row) <= 8 or pd.isna(row.iloc[8]) else str(row.iloc[8]).strip()
            qty_val = f"{g_val}{h_val}"  # 新索引 7+8

            po_val = row.iloc[10] if len(row) > 10 else ""  # 新索引
            lot_val = row.iloc[15] if len(row) > 15 else ""  # 新索引
            rea_val = row.iloc[16] if len(row) > 16 else ""  # 新索引

            sap_key = str(sap_val).strip().lower()
            prod_a, prod_d = db_map.get(sap_key, ("(未找到)", "(未找到)"))

            records.append({
                "DateReceived": date_val,
                "PO#": po_val,
                "Item": item_val,
                "Description": desc_val,
                "Qty": qty_val,
                "SAP": sap_val,
                "Lot": lot_val,
                "Rea": rea_val,
                "ProductA": prod_a,
                "ProductD": prod_d
            })

        df_out = pd.DataFrame(records)
        df_out.to_excel(out_path, index=False)
        messagebox.showinfo("成功", f"✅ 已导出 {len(df_out)} 条记录到：\n{out_path}")

    except Exception as e:
        messagebox.showerror("运行错误", str(e))


# ========== GUI 交互逻辑 ==========
def choose_main_file():
    path = filedialog.askopenfilename(title="选择主文件", filetypes=[("Excel文件", "*.xlsx *.xlsm")])
    if path:
        entry_main.delete(0, tk.END)
        entry_main.insert(0, path)


def choose_db_file():
    path = filedialog.askopenfilename(title="选择数据库", filetypes=[("Excel文件", "*.xlsx *.xlsm")])
    if path:
        entry_db.delete(0, tk.END)
        entry_db.insert(0, path)


def choose_output_file():
    path = filedialog.asksaveasfilename(title="保存输出", defaultextension=".xlsx",
                                        filetypes=[("Excel文件", "*.xlsx"), ("带宏文件", "*.xlsm")])
    if path:
        entry_out.delete(0, tk.END)
        entry_out.insert(0, path)


def on_run():
    main_path = entry_main.get().strip()
    db_path = entry_db.get().strip()
    out_path = entry_out.get().strip()
    input_date = entry_date.get().strip()

    if not (main_path and db_path and out_path and input_date):
        messagebox.showwarning("缺少信息", "请填写所有选项。")
        return
    run_extraction(main_path, db_path, out_path, input_date)


# ========== 修复后的窗口布局 ==========
root = tk.Tk()
root.title("Excel 数据提取工具 v1.2")
root.geometry("600x320")  # 稍微调宽了一点
root.resizable(False, False)

# 使用 padding 让界面不那么拥挤
pad_x = 10
pad_y = 8

# 第一行：主文件
tk.Label(root, text="主文件路径 (MIGO)：").grid(row=0, column=0, sticky="e", pady=pad_y, padx=pad_x)
entry_main = tk.Entry(root, width=50)
entry_main.grid(row=0, column=1, pady=pad_y)
tk.Button(root, text=" 选择 ", command=choose_main_file).grid(row=0, column=2, padx=pad_x)

# 第二行：数据库
tk.Label(root, text="数据库文件 (DB)：").grid(row=1, column=0, sticky="e", pady=pad_y, padx=pad_x)
entry_db = tk.Entry(root, width=50)
entry_db.grid(row=1, column=1, pady=pad_y)
tk.Button(root, text=" 选择 ", command=choose_db_file).grid(row=1, column=2, padx=pad_x)

# 第三行：输出路径
tk.Label(root, text="输出结果路径：").grid(row=2, column=0, sticky="e", pady=pad_y, padx=pad_x)
entry_out = tk.Entry(root, width=50)
entry_out.grid(row=2, column=1, pady=pad_y)
tk.Button(root, text="保存到", command=choose_output_file).grid(row=2, column=2, padx=pad_x)

# 第四行：查询日期
tk.Label(root, text="查询日期：").grid(row=3, column=0, sticky="e", pady=pad_y, padx=pad_x)
frame_date = tk.Frame(root)  # 用个小 frame 把输入框和提示语放一起
frame_date.grid(row=3, column=1, sticky="w")

entry_date = tk.Entry(frame_date, width=15)
entry_date.pack(side="left")
tk.Label(frame_date, text="  (格式: 2024-01-01)", fg="gray").pack(side="left")

# 第五行：运行按钮
tk.Button(root, text=" 🚀 开始提取并匹配数据 ", bg="#4CAF50", fg="white",
          font=("Microsoft YaHei", 11, "bold"), relief="raised",
          command=on_run).grid(row=5, column=0, columnspan=3, pady=25)

root.mainloop()