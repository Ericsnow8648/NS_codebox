import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
from pathlib import Path
import re
from decimal import Decimal
import datetime  # ★ 新增：用于解析 Statement 日期
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter.scrolledtext import ScrolledText


# ===================== Shopee 部分 =====================

def shopee_income_to_csv(income_xlsx: str, output_csv: str):
    """
    从 Shopee 各国已拨款收入表生成统一 CSV（保留两位小数）

    income_xlsx: *.income.已拨款.*.xlsx
    output_csv : 输出 csv 完整路径
    """
    income_path = Path(income_xlsx)

    # 1. 读取 Income 工作表（不设表头）
    df_raw = pd.read_excel(income_path, sheet_name="Income", header=None)

    # 2. 找到标题行（第一列为“编号”的行）
    header_idx_list = df_raw.index[df_raw[0] == "编号"].tolist()
    if not header_idx_list:
        raise ValueError(f"{income_xlsx}: 找不到标题行“编号”。")
    header_idx = header_idx_list[0]

    # 3. 设置表头
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = df_raw.iloc[header_idx]
    df = df[df["订单编号"].notna()].reset_index(drop=True)

    out = pd.DataFrame()

    # 编号
    out["编号"] = pd.to_numeric(df["编号"], errors="coerce").fillna(0).astype(int)

    # 订单编号
    out["订单编号"] = df["订单编号"].astype(str)

    # 拨款完成日期
    dt = pd.to_datetime(df["拨款完成日期"], errors="coerce")
    out["拨款完成日期"] = dt.dt.strftime("%Y/%m/%d")

    # 退款列（兼容：退款金额 / 退款金額）
    refund_col = None
    for c in ["退款金额", "退款金額"]:
        if c in df.columns:
            refund_col = c
            break
    if refund_col is None:
        raise KeyError(f"{income_xlsx}: 找不到退款金额列。")

    # 数值列转成 float
    for col in ["商品原价", "商品折扣", refund_col]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # 付款金额 = 商品原价 + 商品折扣 + 退款金额
    out["付款金额"] = (df["商品原价"] + df["商品折扣"] + df[refund_col]).round(2)

    # 退款金额
    out["退款金额"] = pd.to_numeric(df[refund_col], errors="coerce").fillna(0.0).round(2)

    # 账单金额 = 从 Shopee回扣金额 到 “拨款金额(…币种)” 前一列
    if "Shopee回扣金额" not in df.columns:
        raise KeyError(f"{income_xlsx}: 找不到 'Shopee回扣金额' 列。")

    start_idx = df.columns.get_loc("Shopee回扣金额")

    # 优先找“拨款金额”
    payout_idx = None
    for i, name in enumerate(df.columns):
        if isinstance(name, str) and "拨款金额" in name:
            payout_idx = i
            break

    if payout_idx is not None and payout_idx > start_idx:
        end_idx = payout_idx - 1
    else:
        # 退而求其次：找 Escrow
        escrow_idx = None
        for i, name in enumerate(df.columns):
            if isinstance(name, str) and "Escrow" in name:
                escrow_idx = i
                break
        if escrow_idx is not None and escrow_idx > start_idx:
            end_idx = escrow_idx - 1
        else:
            end_idx = len(df.columns) - 1

    fee_cols = df.columns[start_idx:end_idx + 1]
    fee_sum = df[fee_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1)
    out["账单金额"] = fee_sum.round(2)

    # 过滤：编号>0 & 订单编号非空
    out = out[(out["编号"] > 0) & out["订单编号"].ne("")].reset_index(drop=True)

    out.to_csv(output_csv, index=False, encoding="utf-8-sig")
    return out


def batch_shopee_recursive(root_folder: str, log_func=print):
    """
    从母文件夹递归扫描所有 *.income.已拨款.*.xlsx 并生成 shopee-??1-YYYY-MMDD.csv
    """
    root = Path(root_folder)
    excel_files = sorted(root.rglob("*.income.已拨款*.xlsx"))

    if not excel_files:
        log_func(f"在 {root_folder} 及其子文件夹下没有找到 *.income.已拨款*.xlsx 文件。\n")
        return

    country_map = {
        "PH": "PH", "菲律宾": "PH",
        "TW": "TW", "台湾": "TW",
        "SG": "SG", "シンガポール": "SG",
        "MY": "MY", "马来西亚": "MY",
        "TH": "TH", "泰": "TH",
        "BR": "BR", "ブラジル": "BR",
        "VN": "VN", "ベトナム": "VN",
        "ID": "ID", "インドネシア": "ID",
    }

    for xlsx in excel_files:
        if xlsx.name.startswith("~$"):
            continue

        folder_name = xlsx.parent.name

        # 店铺编号
        m_store = re.search(r"(C\d{6})", folder_name)
        store_code = m_store.group(1) if m_store else "STORE"

        # 国家代码：优先看文件名中的 .xx.，否则看文件夹名
        stem = xlsx.stem
        m_cc = re.search(r"\.([a-z]{2})\.", stem)
        if m_cc:
            cc = m_cc.group(1).upper()
        else:
            cc = "XX"
            upper = folder_name.upper()
            for key, val in country_map.items():
                if key.upper() in upper:
                    cc = val
                    break

        # 输出文件名：shopee-CC-STORE-YYYY-MMDD.csv
        m_date = re.search(r"(\d{8})", stem)
        if m_date:
            yyyymmdd = m_date.group(1)
            yyyy, mm, dd = yyyymmdd[:4], yyyymmdd[4:6], yyyymmdd[6:8]
        else:
            yyyy, mm, dd = "0000", "00", "00"
        out_name = f"shopee-{cc}-{store_code}-{yyyy}-{mm}{dd}.csv"
        out_path = xlsx.with_name(out_name)

        try:
            df_out = shopee_income_to_csv(xlsx, out_path)
            rel_in = xlsx.relative_to(root)
            rel_out = out_path.relative_to(root)
            log_func(f"[Shopee OK] {rel_in} -> {rel_out} ({len(df_out)} 行)\n")
        except Exception as e:
            rel_in = xlsx.relative_to(root)
            log_func(f"[Shopee ERROR] {rel_in}: {e}\n")


# ===================== Lazada 部分 =====================

def safe_order_no(x):
    """
    Lazada 订单号安全转字符串：
    - 去掉科学计数法
    - 去掉末尾的 .0（由 float 转换时产生）
    - 保留所有位数，避免前导零丢失
    """
    if pd.isna(x):
        return ""
    s = str(x).strip()

    # 科学计数法，如 1.234E+17
    if "e" in s.lower():
        s = format(Decimal(s), "f")

    # float 整数形态，如 '123456789012345678.0'
    if s.endswith(".0"):
        s = s[:-2]

    return s


def _get_payout_date_from_statement_or_txn(df: pd.DataFrame) -> datetime.date:
    """
    先尝试从 Statement 列提取结算周期结束日，例如：
    '17 Nov 2025 - 23 Nov 2025' → 2025-11-23（date）

    如果没有 Statement 或解析失败，则退回：
    Transaction Date 最大值（原有逻辑）
    """
    # 优先使用 Statement
    if "Statement" in df.columns:
        stmt_series = df["Statement"].dropna()
        if not stmt_series.empty:
            stmt = str(stmt_series.iloc[0]).strip()
            # 找到所有形如 '23 Nov 2025' 的片段，取最后一个（结束日期）
            matches = re.findall(r"(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})", stmt)
            if matches:
                last_date_str = matches[-1]
                try:
                    dt = datetime.datetime.strptime(last_date_str, "%d %b %Y").date()
                    return dt
                except Exception:
                    pass  # 解析失败就走 fallback

    # fallback：Transaction Date 最大值
    tdates = pd.to_datetime(df.get("Transaction Date"), errors="coerce")
    if tdates.notna().any():
        return tdates.max().date()

    # 再兜底：用今天（一般不会走到这里，除非文件几乎没数据）
    return datetime.date.today()


def lazada_from_transaction(trans_xlsx: str,
                             country_code: str,
                             store_code: str,
                             output_csv: str | None = None):
    """
    通用 Lazada 已拨款处理（基于 Transaction Overview）

    - 一行 = 一个 Order No.
    - 不过滤 Order No == 0（你要求保留）
    - 付款金额 = Item Price Credit + Seller Virtual Credit - Co-fund Price Cut 合计
    - 退款金額 = 0（将来有退款再扩展）
    - 账单金额 = 所有 Amount 合计 - 付款金额
    - 拨款完成日期：优先从 Statement 中取结算周期结束日期，
      若失败则回退为 Transaction Date 最大值
    """
    trans_path = Path(trans_xlsx)

    df = pd.read_excel(trans_path, sheet_name="Transaction Overview", header=0)
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)

    # ★★ 拨款完成日期：先看 Statement，再看 Transaction Date
    payout_date = _get_payout_date_from_statement_or_txn(df)
    payout_date_str = f"{payout_date.year}/{payout_date.month}/{payout_date.day}"

    rows = []

    for order_no, g in df.groupby("Order No."):
        if pd.isna(order_no):
            continue  # 只跳过 NaN

        # ★★★ 修改点：付款金额 = Item Price Credit + Seller Virtual Credit - Co-fund Price Cut
        pay_item = g.loc[g["Fee Name"] == "Item Price Credit", "Amount"].sum()
        pay_cofund = g.loc[g["Fee Name"] == "Seller Virtual Credit - Co-fund Price Cut", "Amount"].sum()
        pay = pay_item + pay_cofund

        refund = 0.0
        total = g["Amount"].sum()
        bill = total - pay  # 账单金额依然是：所有 Amount 合计 - 付款金额

        rows.append([
            safe_order_no(order_no),   # 订单编号强制为“安全字符串”
            round(pay, 2),
            round(refund, 2),
            round(bill, 2),
        ])

    out = pd.DataFrame(rows, columns=["订单编号", "付款金额", "退款金額", "账单金额"])
    out = out.sort_values("订单编号").reset_index(drop=True)
    out.insert(0, "编号", range(1, len(out) + 1))
    out["拨款完成日期"] = payout_date_str

    # 再次强制为字符串，防止后面被推断为数字类型
    out["订单编号"] = out["订单编号"].astype(str)

    out = out[["编号", "订单编号", "拨款完成日期", "付款金额", "退款金額", "账单金额"]]

    # 输出文件名：Lazada-CC-STORE-YYYY-MMDD.csv
    if output_csv is None:
        m = re.search(r"(\d{8})", trans_path.stem)
        if m:
            yyyymmdd = m.group(1)
            yyyy, mm, dd = yyyymmdd[:4], yyyymmdd[4:6], yyyymmdd[6:8]
        else:
            yyyy, mm, dd = (
                str(payout_date.year),
                f"{payout_date.month:02d}",
                f"{payout_date.day:02d}",
            )
        out_name = f"Lazada-{country_code}-{store_code}-{yyyy}-{mm}{dd}.csv"
        out_path = trans_path.with_name(out_name)
    else:
        out_path = Path(output_csv)

    out.to_csv(out_path, index=False, encoding="utf-8-sig")
    return out, out_path


def batch_lazada_all(root_folder: str, log_func=print):
    """
    递归扫描母文件夹，处理所有 Lazada *.已拨款*.xlsx
    """
    root = Path(root_folder)
    files = sorted(root.rglob("*.已拨款*.xlsx"))

    if not files:
        log_func(f"在 {root_folder} 及其子文件夹下没有找到 Lazada *.已拨款*.xlsx 文件。\n")
        return

    country_map = {
        "PH": "PH", "フィリピン": "PH",
        "MY": "MY", "マレーシア": "MY",
        "TH": "TH", "タイ": "TH",
        "SG": "SG", "シンガポール": "SG",
        "ID": "ID", "インドネシア": "ID",
        "VN": "VN", "ベトナム": "VN",
    }

    for f in files:
        if f.name.startswith("~$"):
            continue

        folder_name = f.parent.name

        # 店铺编号
        m_store = re.search(r"(C\d{6})", folder_name)
        store_code = m_store.group(1) if m_store else "STORE"

        # 国家代码：优先从文件名 .xx. 抽取，再从文件夹名关键词
        stem = f.stem
        m_cc = re.search(r"\.([a-z]{2})\.", stem)
        if m_cc:
            cc = m_cc.group(1).upper()
        else:
            cc = "XX"
            upper = folder_name.upper()
            for key, val in country_map.items():
                if key.upper() in upper:
                    cc = val
                    break

        try:
            df_out, out_path = lazada_from_transaction(f, cc, store_code)
            rel_in = f.relative_to(root)
            rel_out = out_path.relative_to(root)
            log_func(f"[Lazada OK] {rel_in} -> {rel_out} ({len(df_out)} 行)\n")
        except Exception as e:
            rel_in = f.relative_to(root)
            log_func(f"[Lazada ERROR] {rel_in}: {e}\n")


# ===================== Tkinter GUI =====================

def main():
    DEFAULT_SHOPEE = r"C:\Users\mitsu\OneDrive\デスクトップ\shopee"
    DEFAULT_LAZADA = r"C:\Users\mitsu\OneDrive\デスクトップ\lazada"

    root = tk.Tk()
    root.title("Shopee / Lazada 已拨款 数据批量处理")

    # ------- Shopee 路径 -------
    frm_shopee = ttk.LabelFrame(root, text="Shopee 母文件夹", padding=10)
    frm_shopee.pack(fill="x", padx=10, pady=(10, 5))

    shopee_var = tk.StringVar(value=DEFAULT_SHOPEE)
    ttk.Label(frm_shopee, text="路径:").pack(side="left")
    ent_shopee = ttk.Entry(frm_shopee, textvariable=shopee_var, width=80)
    ent_shopee.pack(side="left", padx=5, fill="x", expand=True)

    def choose_shopee():
        folder = filedialog.askdirectory(initialdir=shopee_var.get() or DEFAULT_SHOPEE)
        if folder:
            shopee_var.set(folder)

    ttk.Button(frm_shopee, text="选择...", command=choose_shopee).pack(side="left", padx=5)

    # ------- Lazada 路径 -------
    frm_lazada = ttk.LabelFrame(root, text="Lazada 母文件夹", padding=10)
    frm_lazada.pack(fill="x", padx=10, pady=5)

    lazada_var = tk.StringVar(value=DEFAULT_LAZADA)
    ttk.Label(frm_lazada, text="路径:").pack(side="left")
    ent_lazada = ttk.Entry(frm_lazada, textvariable=lazada_var, width=80)
    ent_lazada.pack(side="left", padx=5, fill="x", expand=True)

    def choose_lazada():
        folder = filedialog.askdirectory(initialdir=lazada_var.get() or DEFAULT_LAZADA)
        if folder:
            lazada_var.set(folder)

    ttk.Button(frm_lazada, text="选择...", command=choose_lazada).pack(side="left", padx=5)

    # ------- 按钮 & 日志 -------
    frm_btn = ttk.Frame(root, padding=(10, 0, 10, 5))
    frm_btn.pack(fill="x")

    log_box = ScrolledText(root, height=20)
    log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def append_log(msg: str):
        log_box.insert("end", msg)
        log_box.see("end")
        root.update_idletasks()

    def run_shopee():
        folder = shopee_var.get().strip()
        if not folder:
            append_log("请先选择 Shopee 母文件夹。\n")
            return
        append_log(f"=== 开始处理 Shopee: {folder} ===\n")
        batch_shopee_recursive(folder, log_func=append_log)
        append_log("=== Shopee 处理结束 ===\n\n")

    def run_lazada():
        folder = lazada_var.get().strip()
        if not folder:
            append_log("请先选择 Lazada 母文件夹。\n")
            return
        append_log(f"=== 开始处理 Lazada: {folder} ===\n")
        batch_lazada_all(folder, log_func=append_log)
        append_log("=== Lazada 处理结束 ===\n\n")

    btn_shopee = ttk.Button(frm_btn, text="处理 Shopee", command=run_shopee)
    btn_shopee.pack(side="left", padx=5)
    btn_lazada = ttk.Button(frm_btn, text="处理 Lazada", command=run_lazada)
    btn_lazada.pack(side="left", padx=5)

    root.mainloop()


if __name__ == "__main__":
    main()
