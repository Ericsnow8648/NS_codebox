# -*- coding: utf-8 -*-
"""
Shopee + Lazada PDF → Payoneer CSV 自动填充程序
------------------------------------------------
功能概要：
1. 扫描当前目录及子目录下所有 PDF（不限文件名）
2. 自动识别 PDF 内容是 Shopee 还是 Lazada
3. Shopee：
   - 从「总结支出」区域解析：
       ・売上：买家支付的商品金额
       ・入金：总拨款金额（当地币）
       ・汇率：PDF 上的汇率
       ・USD 金额：总拨款金额（USD）
   - 解析 "Statement for 2025-04-23" 作为结算日期
   - 在 Payoneer CSV 中按以下规则选一行填入：
       a) Description 中包含 Shopee + 国家关键字
       b) Currency = USD
       c) Amount 与 PDF USD 金额精确匹配（误差 < 0.01）
       d) 若有多行金额相同：
          - 若 USD + Description + Date 三者都相同且≥2行 → 视为无法区分，留空并记日志
          - 否则选与结算日期最接近的“完全空白行”写入
4. Lazada：
   - 从 PDF 中的「货款」「Total Settlement」解析卖家收款和拨款金额
   - 先按币种对周报排序，对「折算 USD < 1」的周期与之后同币种周期合并：
       ・使用 EXPECTED_RATE[currency] 粗略折算 USD
       ・若单周 <1 USD，则依次与下一周累加 Total Settlement，直到合并后的 approx_usd ≥1
       ・合并周期的截止日期 end_date 取最后一个周报的 end_date
   - 匹配 Payoneer CSV 时：
       a) 只看 Currency=USD 的行（优先 Description 含 Lazada）
       b) 若有 end_date：CSV.Date 必须满足  end_date ≤ Date ≤ end_date+5天
       c) 计算隐含汇率 total_local / Amount_USD，要求落在币种固定区间 RATE_RANGE 内
       d) 在候选中选隐含汇率最接近 EXPECTED_RATE 的一行
   - 每个 Lazada PDF（或合并后的周期）只匹配一个 CSV 行，并写入 LazadaCountry 列（PH/MY/TH/SG/VN）
5. 只填“完全空白行”，不会覆盖人工修改
6. 所有金额写入 CSV 时统一为两位小数的字符串（内部计算仍用 float）
7. 输出：
   - 原 CSV 的镜像目录结构，文件名加 `_filled`
   - 日志 CSV（记录未匹配等情况、Shopee 无法区分等）
"""

import re
import csv
from pathlib import Path
from datetime import datetime, date

import pdfplumber
import pandas as pd


# ========= 配置区域 =========
ROOT_DIR = Path(".").resolve()          # 脚本根目录
OUTPUT_ROOT = ROOT_DIR / "output_filled"
OUTPUT_SUFFIX = "_filled"
# ==========================


# ================= 工具函数 =================

def to_float_safe(v):
    """安全转换为 float，失败返回 None。"""
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.replace(",", "").strip()
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def fmt2(v):
    """
    金额格式化：保留两位小数的字符串。
    内部计算仍用 float，仅在写到 CSV 时控制展示。
    """
    x = to_float_safe(v)
    if x is None:
        return None
    return f"{x:.2f}"


def is_blank_value(v):
    """判断一个值是否视为“空”"""
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def row_blank_for_fill(row, cols):
    """
    判断一行是否“完全空白”：
    只要目标列里有任意一个非空，就视为已填过，不再覆盖。
    """
    for c in cols:
        if not is_blank_value(row.get(c)):
            return False
    return True


# ================= PDF 解析 =================

def extract_numbers(text: str):
    """从文本中提取所有数字（支持负号、千分位、小数）"""
    return [float(n.replace(",", "")) for n in re.findall(r"-?\d[\d,]*\.?\d*", text)]


def parse_lazada_date(text: str):
    """
    Lazada PDF 日期区间：
    例如: "10/3/2025 to 16/3/2025"
    返回 (start_date, end_date) 或 (None, None)
    """
    m = re.search(
        r"(\d{1,2}/\d{1,2}/\d{4})\s*(?:to|-)\s*(\d{1,2}/\d{1,2}/\d{4})",
        text
    )
    if not m:
        return None, None
    try:
        start = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        end = datetime.strptime(m.group(2), "%d/%m/%Y").date()
        return start, end
    except ValueError:
        return None, None


def parse_shopee_statement_date(text: str):
    """
    Shopee PDF 结算日期：
    典型格式：
      - "Statement for 2025-04-23"
      - "Statement for 23/04/2025"
    返回 date 或 None。
    """
    m1 = re.search(r"Statement\s+for\s+(\d{4}-\d{2}-\d{2})", text)
    if m1:
        try:
            return datetime.strptime(m1.group(1), "%Y-%m-%d").date()
        except ValueError:
            pass

    m2 = re.search(r"Statement\s+for\s+(\d{1,2}/\d{1,2}/\d{4})", text)
    if m2:
        try:
            return datetime.strptime(m2.group(1), "%d/%m/%Y").date()
        except ValueError:
            pass

    return None


def parse_pdf(pdf_path: Path):
    """
    自动识别 PDF 是 Shopee 还是 Lazada，并解析核心金额。

    返回：
    {
        "path": Path,
        "type": "shopee" / "lazada",
        "currency": "PHP"/"BRL"/...,
        "sale": 卖家收款（Shopee=买家支付金额，Lazada=货款）,
        "total_local": 拨款金额（当地币）,
        "rate": Shopee=PDF汇率, Lazada=None,
        "usd": Shopee=USD拨款金额, Lazada=None,
        "end_date": 结算日/区间结束日（date）或 None
    }
    """
    with pdfplumber.open(str(pdf_path)) as pdf:
        texts = [page.extract_text() or "" for page in pdf.pages]

    full_text = "\n".join(texts)

    # 币种：兼容「金额 (BRL)」「Amount (PHP)」
    m_cur = re.search(r"(?:金额|Amount)\s*\(([A-Z]{3})\)", full_text)
    currency = m_cur.group(1) if m_cur else None

    # ---------- Lazada ----------
    if "Total Settlement" in full_text:
        # 尝试找“货款”
        m_sale = re.search(r"货款\s*([-\d,\.]+)", full_text)
        if m_sale:
            sale = float(m_sale.group(1).replace(",", ""))
        else:
            # 没有货款字段（纯费用 / 负数周报），视为卖上=0
            sale = 0.0

        m_total = re.search(r"Total\s+Settlement\s+([-\d,\.]+)", full_text)
        if not m_total:
            raise ValueError(f"{pdf_path.name}: 未找到『Total Settlement』金额")
        total_local = float(m_total.group(1).replace(",", ""))

        if currency is None:
            raise ValueError(f"{pdf_path.name}: Lazada 无法识别币种（Amount (XXX)）")

        _, end_date = parse_lazada_date(full_text)

        return {
            "path": pdf_path,
            "type": "lazada",
            "currency": currency,
            "sale": sale,
            "total_local": total_local,
            "rate": None,
            "usd": None,
            "end_date": end_date,
        }

    # ---------- Shopee ----------
    if "总结支出" in full_text:
        try:
            start = full_text.index("总结支出")
        except ValueError:
            start = 0
        try:
            end = full_text.index("**详细调整内容")
        except ValueError:
            end = len(full_text)

        sub = full_text[start:end]
        nums = extract_numbers(sub)
        if len(nums) < 4:
            raise ValueError(f"{pdf_path.name}: Shopee 总结区数字过少: {nums}")

        sale = nums[0]
        total_local = nums[-3]
        rate = nums[-2]
        usd = nums[-1]

        if currency is None:
            raise ValueError(f"{pdf_path.name}: Shopee 无法识别币种（金额 (XXX)）")

        end_date = parse_shopee_statement_date(full_text)

        return {
            "path": pdf_path,
            "type": "shopee",
            "currency": currency,
            "sale": sale,
            "total_local": total_local,
            "rate": rate,
            "usd": usd,
            "end_date": end_date,
        }

    raise ValueError(f"{pdf_path.name}: 无法识别为 Shopee 或 Lazada 格式（缺少关键字）")


def load_all_pdfs():
    pdf_files = sorted(ROOT_DIR.rglob("*.pdf"))
    parsed_list = []
    for p in pdf_files:
        try:
            d = parse_pdf(p)
            parsed_list.append(d)
            print(
                f"[PDF] {p.name} type={d['type']} "
                f"currency={d['currency']} sale={d['sale']} "
                f"local={d['total_local']} end={d['end_date']}"
            )
        except Exception as e:
            print(f"[PDF-SKIP] {p.name}: {e}")
    return parsed_list


# ================= 日期解析 =================

def parse_csv_date(s: str):
    """将 CSV 中的 Date 字符串解析为 date 对象。"""
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None

    fmts = ("%d %b, %Y", "%d %b %Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y")
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    try:
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None


# ================= Lazada 匹配 =================

# 预期汇率（当地币 / USD），用于在多个候选中评分（越接近越好）
EXPECTED_RATE = {
    "PHP": 57.0,
    "MYR": 4.5,
    "THB": 33.0,
    "SGD": 1.32,
    "VND": 26000.0,
}

# 固定允许的隐含汇率区间（当地币 / USD）
RATE_RANGE = {
    "VND": (25000.0, 27000.0),
    "PHP": (54.0, 60.0),
    "THB": (28.0, 36.0),
    "MYR": (4.1, 5.0),
    "SGD": (1.25, 1.38),
}

# Lazada 币种 -> 国家代码
CURRENCY_TO_COUNTRY = {
    "PHP": "PH",
    "MYR": "MY",
    "THB": "TH",
    "SGD": "SG",
    "VND": "VN",
}


def merge_small_lazada_pdfs(parsed_list, usd_threshold=1.0):
    """
    将 Lazada 中「折算 USD < usd_threshold」的周期，与后面同币种的周期合并。
    合并规则：
      - 使用 EXPECTED_RATE[currency] 估算 USD
      - 若 approx_usd < usd_threshold，则依次把后面同币种的周期加总，
        直到合并后的 approx_usd >= usd_threshold，或没有更多同币种周期
      - end_date 采用最后一个周期的 end_date
    返回：新的 PDF 列表（Shopee 原样保留，Lazada 替换为合并后的周期）
    """

    shopee = [d for d in parsed_list if d["type"] == "shopee"]
    lazada = [d for d in parsed_list if d["type"] == "lazada"]

    lazada_sorted = sorted(
        lazada,
        key=lambda d: (
            d.get("currency") or "",
            d.get("end_date") or date.min,
            d["path"].name,
        )
    )

    merged_lazada = []
    i = 0

    while i < len(lazada_sorted):
        cur = lazada_sorted[i]
        currency = cur.get("currency")
        rate_est = EXPECTED_RATE.get(currency)

        total_local = to_float_safe(cur.get("total_local"))
        sale = to_float_safe(cur.get("sale"))
        end_date = cur.get("end_date")

        if rate_est is None or total_local is None:
            merged_lazada.append(cur)
            i += 1
            continue

        approx_usd = total_local / rate_est

        if approx_usd >= usd_threshold:
            merged_lazada.append(cur)
            i += 1
            continue

        # 需要与后续同币种周期合并
        j = i + 1
        merged_from = [cur["path"].name]

        while j < len(lazada_sorted):
            nxt = lazada_sorted[j]
            if nxt.get("currency") != currency:
                break

            tl_next = to_float_safe(nxt.get("total_local"))
            sale_next = to_float_safe(nxt.get("sale"))

            if tl_next is not None:
                total_local += tl_next
            if sale_next is not None:
                sale = (sale or 0) + sale_next

            if nxt.get("end_date") is not None:
                end_date = nxt["end_date"]

            merged_from.append(nxt["path"].name)

            approx_usd = total_local / rate_est
            j += 1

            if approx_usd >= usd_threshold:
                break

        new_entry = cur.copy()
        new_entry["sale"] = sale
        new_entry["total_local"] = total_local
        new_entry["end_date"] = end_date
        new_entry["merged_from"] = merged_from

        merged_lazada.append(new_entry)
        i = j

    return shopee + merged_lazada


def find_best_lazada_row(
    df,
    pdf_info,
    used_idx: set,
    max_future_days=10,
):
    """
    Lazada 匹配逻辑（使用固定汇率范围 + 单向日期约束）：
      1) 只考虑 Currency = USD 的行（优先 Description 含 Lazada）
      2) 如果 PDF 有 end_date，则要求：
           CSV.Date >= end_date 且 CSV.Date - end_date <= max_future_days
      3) 计算隐含汇率 implied = total_local / abs(Amount_USD)
         只有当 implied 位于 RATE_RANGE[currency] 区间内时才认为是候选
      4) 在所有候选中选「隐含汇率最接近 EXPECTED_RATE[currency]」的一行
    """
    currency = pdf_info["currency"]
    total_local = pdf_info["total_local"]
    end_date = pdf_info.get("end_date")

    expected_rate = EXPECTED_RATE.get(currency)
    rate_range = RATE_RANGE.get(currency)
    if not expected_rate or not rate_range:
        return None
    min_rate, max_rate = rate_range

    if not {"Description", "Currency", "Amount", "Date"}.issubset(df.columns):
        return None

    desc_s = df["Description"].fillna("").astype(str)
    curr_s = df["Currency"].fillna("").astype(str)
    date_s = df["Date"].fillna("").astype(str)

    # 先选择 Currency=USD 且描述里有 Lazada 的行
    mask = curr_s.eq("USD") & desc_s.str.contains("Lazada", case=False, na=False)
    candidates = df.index[mask].tolist()
    if not candidates:
        candidates = df.index[curr_s.eq("USD")].tolist()

    best_idx = None
    best_diff = None

    for idx in candidates:
        if idx in used_idx:
            continue

        # 日期过滤：只允许 [end_date, end_date + max_future_days]
        if end_date is not None:
            dt = parse_csv_date(date_s.at[idx])
            if dt is None:
                continue
            if dt < end_date:
                continue
            if (dt - end_date).days > max_future_days:
                continue

        amt_usd = to_float_safe(df.at[idx, "Amount"])
        if not amt_usd:
            continue

        implied_rate = total_local / abs(amt_usd)

        # 固定区间过滤
        if implied_rate < min_rate or implied_rate > max_rate:
            continue

        diff = abs(implied_rate - expected_rate)

        if best_diff is None or diff < best_diff:
            best_diff = diff
            best_idx = idx

    return best_idx


# ================= CSV 读取 =================

def load_csv_tables():
    csv_files = sorted(ROOT_DIR.rglob("*.csv"))
    tables = []
    for p in csv_files:
        if OUTPUT_ROOT in p.parents:
            continue
        try:
            df = pd.read_csv(p, dtype=str)
            tables.append({"path": p, "df": df})
            print(f"[CSV] 读取：{p.relative_to(ROOT_DIR)} ({len(df)} 行)")
        except Exception as e:
            print(f"[CSV-ERR] {p.name}: {e}")
    return tables


# ================= 主流程 =================

def main():
    parsed_pdfs = load_all_pdfs()
    if not parsed_pdfs:
        print("⚠ 未找到任何 PDF")
        return

    # Lazada 小额周期合并
    parsed_pdfs = merge_small_lazada_pdfs(parsed_pdfs, usd_threshold=1.0)

    tables = load_csv_tables()
    if not tables:
        print("⚠ 未找到任何 CSV")
        return

    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    log_rows = []
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = OUTPUT_ROOT / f"log_{ts}.csv"

    REQUIRED_BASE = ["Date", "Description", "Amount", "Currency"]
    TARGET_COLS = [
        "売上",
        "手数料",
        "入金",
        "汇率",
        "表格中汇后金额",
        "计算汇后金额",
        "验证",
        "LazadaCountry",
    ]

    # Shopee：币种 -> Description 关键字
    SHOPEE_DESC_KEY = {
        "BRL": "Shopee Brazil",
        "TWD": "Shopee Taiwan",
        "THB": "Shopee Thailand",
        "PHP": "Shopee- Philippines",
        "MYR": "Shopee- Malaysia",
        "VND": "Shopee VN",
        "SGD": "ShopeePay Singapore USD",
    }

    # Lazada：防止一个 CSV 中多次使用同一行
    lazada_used_by_dfid = {}
    lazada_pdf_filled_once = set()

    for pdf in parsed_pdfs:
        pdf_name = pdf["path"].name
        pdf_type = pdf["type"]
        currency = pdf["currency"]

        print(f"\n===== 处理 PDF: {pdf_name} ({pdf_type}, {currency}) =====")

        matched = False

        for t in tables:
            df = t["df"]
            path = t["path"]
            df_id = id(df)

            if not all(c in df.columns for c in REQUIRED_BASE):
                continue

            for col in TARGET_COLS:
                if col not in df.columns:
                    df[col] = None

            # ---------- Shopee ----------
            if pdf_type == "shopee":
                sale = pdf["sale"]
                total_local = pdf["total_local"]
                rate = pdf["rate"]
                usd = pdf["usd"]
                end_date = pdf.get("end_date")

                desc_s = df["Description"].fillna("").astype(str)
                curr_s = df["Currency"].fillna("").astype(str)
                date_s = df["Date"].fillna("").astype(str)

                key = SHOPEE_DESC_KEY.get(currency, "Shopee")

                mask = curr_s.eq("USD") & \
                       desc_s.str.contains("Shopee", na=False) & \
                       desc_s.str.contains(key, na=False)

                candidates = df.index[mask].tolist()
                if not candidates:
                    continue

                amount_matches = []
                for idx in candidates:
                    amt = to_float_safe(df.at[idx, "Amount"])
                    if amt is None:
                        continue
                    if abs(amt - usd) < 0.01:
                        amount_matches.append(idx)

                if not amount_matches:
                    continue

                blank_candidates = []
                nonblank_matches = []

                for idx in amount_matches:
                    row = df.loc[idx]
                    if row_blank_for_fill(row, TARGET_COLS):
                        blank_candidates.append(idx)
                    else:
                        nonblank_matches.append(idx)

                if not blank_candidates and nonblank_matches:
                    matched = True
                    break

                if not blank_candidates:
                    continue

                desc_set = {desc_s.at[i].strip() for i in blank_candidates}
                date_parsed_set = {parse_csv_date(date_s.at[i]) for i in blank_candidates}
                date_set_no_none = {d for d in date_parsed_set if d is not None}

                if len(blank_candidates) >= 2 and len(desc_set) == 1 and len(date_set_no_none) == 1:
                    msg = f"Shopee {currency}：存在多个 USD/Description/Date 完全相同的空行，未自动填充"
                    print("[AMBIGUOUS]", msg)
                    log_rows.append({
                        "pdf": pdf_name,
                        "action": "ambiguous_same_usd_desc_date",
                        "msg": msg,
                    })
                    matched = True
                    break

                best_idx = None
                best_diff_days = None

                for idx in blank_candidates:
                    if end_date is not None:
                        dt = parse_csv_date(date_s.at[idx])
                        if dt is None:
                            diff = 9999
                        else:
                            diff = abs((dt - end_date).days)
                    else:
                        diff = 0

                    if best_idx is None or diff < best_diff_days:
                        best_idx = idx
                        best_diff_days = diff

                if best_idx is None:
                    best_idx = blank_candidates[0]

                idx = best_idx
                row = df.loc[idx]
                amt = to_float_safe(row["Amount"])
                if amt is None:
                    continue

                fee = total_local - sale
                calc_after = total_local * rate

                df.at[idx, "売上"] = fmt2(sale)
                df.at[idx, "入金"] = fmt2(total_local)
                df.at[idx, "手数料"] = fmt2(fee)
                df.at[idx, "汇率"] = rate
                df.at[idx, "表格中汇后金额"] = fmt2(amt)
                df.at[idx, "计算汇后金额"] = fmt2(calc_after)
                df.at[idx, "验证"] = None if fmt2(amt) == fmt2(calc_after) else "false"

                print(f"[FILL] Shopee → {path.name} 第 {idx + 2} 行")
                matched = True

            # ---------- Lazada ----------
            elif pdf_type == "lazada":
                if pdf_name in lazada_pdf_filled_once:
                    matched = True
                    break

                sale = pdf["sale"]
                total_local = pdf["total_local"]

                used_idx = lazada_used_by_dfid.setdefault(df_id, set())

                idx = find_best_lazada_row(
                    df,
                    pdf_info=pdf,
                    used_idx=used_idx,
                    max_future_days=10
                )
                if idx is None:
                    continue

                row = df.loc[idx]
                if not row_blank_for_fill(row, TARGET_COLS):
                    used_idx.add(idx)
                    matched = True
                    break

                amt_usd = to_float_safe(row["Amount"])
                if not amt_usd:
                    used_idx.add(idx)
                    continue

                rate = amt_usd / total_local
                calc_after = total_local * rate
                fee = total_local - sale

                df.at[idx, "売上"] = fmt2(sale)
                df.at[idx, "入金"] = fmt2(total_local)
                df.at[idx, "手数料"] = fmt2(fee)
                df.at[idx, "汇率"] = rate
                df.at[idx, "表格中汇后金额"] = fmt2(amt_usd)
                df.at[idx, "计算汇后金额"] = fmt2(calc_after)
                df.at[idx, "验证"] = None if fmt2(amt_usd) == fmt2(calc_after) else "false"

                country_code = CURRENCY_TO_COUNTRY.get(currency)
                if country_code:
                    df.at[idx, "LazadaCountry"] = country_code

                used_idx.add(idx)
                lazada_pdf_filled_once.add(pdf_name)

                print(f"[FILL] Lazada → {path.name} 第 {idx + 2} 行 (country={country_code})")
                matched = True

            if matched:
                break

        if not matched:
            print(f"[NOT_FOUND] {pdf_name}: 未匹配 CSV 行")
            log_rows.append({
                "pdf": pdf_name,
                "action": "not_found",
                "msg": f"{pdf_type} {currency} 未匹配 CSV 行",
            })

    # ---------- 保存 _filled CSV ----------
    for t in tables:
        orig = t["path"]
        df = t["df"]
        rel = orig.relative_to(ROOT_DIR)

        out_dir = OUTPUT_ROOT / rel.parent
        out_dir.mkdir(parents=True, exist_ok=True)

        out_path = out_dir / (orig.stem + OUTPUT_SUFFIX + orig.suffix)

        try:
            df.to_csv(out_path, index=False, encoding="utf-8-sig")
            print(f"[SAVE] {rel} -> {out_path.relative_to(ROOT_DIR)}")
        except PermissionError:
            backup_path = out_dir / (orig.stem + OUTPUT_SUFFIX + f"_{ts}" + orig.suffix)
            df.to_csv(backup_path, index=False, encoding="utf-8-sig")
            print(f"[SAVE] {rel} 被占用，改为另存：{backup_path.relative_to(ROOT_DIR)}")

    # ---------- 保存日志 ----------
    with log_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["pdf", "action", "msg"])
        writer.writeheader()
        for r in log_rows:
            writer.writerow(r)

    print("\n📄 日志文件：", log_path)
    print("🎉 完成！所有结果在 output_filled/ 目录中。")


if __name__ == "__main__":
    main()
