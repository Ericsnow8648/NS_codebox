# -*- coding: utf-8 -*-
"""
NetSuite 入金（CustPymt）适用请求書 RPA

规则：
1) 打开指定URL
2) 用户手动登录后按回车继续
3) 读取“请求書列表.xlsx”的「請求書番号」「請求書金額」
4) 每次输入前读取“支払額”
5) 在“アイテム選択(autoenter)”输入請求書番号并回车（触发适用）
6) 再读“支払額”，校验差额是否等于该行「請求書金額」，写日志
7) 循环直到全部适用
8) 适用全部后：回车前再次校验「支払額 = Excel 金额合计」，并高亮提示任意“金额有误”；等待用户回车后结束并关闭浏览器
附加：
- 自动处理 NetSuite 弹窗（alert），并把弹窗内容写入 CSV 的 alert_text 列
"""

import os
import time
import re
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException

# （可选）Windows 控制台彩色输出
try:
    import colorama
    colorama.just_fix_windows_console()
    _COLOR_OK = True
except Exception:
    _COLOR_OK = False


URL = "https://6806569.app.netsuite.com/app/accounting/transactions/custpymt.nl?id=6912959&whence=&e=T&cp=T&memdoc=0"

# ⚠️ 注意：路径末尾不要有空格
BASE_DIR = r"C:\Users\mitsu\OneDrive\デスクトップ\amazon（toB）请求書自动录入".strip()
EXCEL_PATH = os.path.join(BASE_DIR, "请求書列表.xlsx")

# 日志输出位置（同目录）
LOG_DIR = BASE_DIR

# 如果你的 chromedriver 已经在 PATH，可把它设为 None
CHROME_DRIVER_PATH = None  # 例如 r"C:\tools\chromedriver.exe"

# 金额比较允许误差（避免显示格式导致的 0.01 级误差）
TOLERANCE = Decimal("0.01")


def c_red(s: str) -> str:
    if not _COLOR_OK:
        return s
    return f"\033[31m\033[1m{s}\033[0m"  # 红色加粗


def c_yellow(s: str) -> str:
    if not _COLOR_OK:
        return s
    return f"\033[33m\033[1m{s}\033[0m"  # 黄色加粗


def to_decimal_money(x) -> Decimal:
    """
    将页面/Excel中的金额转换为 Decimal(两位小数)
    - 支持 1,234.56 / 1234 / 0.00 / 空 / None
    - 去除千分位、货币符号、空格等
    """
    if x is None:
        return Decimal("0.00")

    if isinstance(x, (int, float, Decimal)):
        s = str(x)
    else:
        s = str(x).strip()

    if s == "" or s.lower() == "nan":
        return Decimal("0.00")

    s = re.sub(r"[^\d\.\-]", "", s)  # 仅保留数字/小数点/负号
    if s in ("", "-", ".", "-."):
        return Decimal("0.00")

    d = Decimal(s)
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def build_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")

    # 兼容旧版 Selenium 写法；如你环境是 Selenium 4+，也可改为 Service 写法
    if CHROME_DRIVER_PATH:
        return webdriver.Chrome(executable_path=CHROME_DRIVER_PATH, options=options)
    return webdriver.Chrome(options=options)


def wait_for_manual_login():
    input("请在浏览器中手动登录 NetSuite。登录完成后回车继续程序：")


def accept_any_alert(driver: webdriver.Chrome, max_rounds: int = 3):
    """
    如存在 alert，自动点击“确定(accept)”。
    返回：(handled: bool, texts: list[str])
    处理连续弹窗（最多 max_rounds 次）
    """
    texts = []
    handled = False
    for _ in range(max_rounds):
        try:
            alert = driver.switch_to.alert
            txt = alert.text or ""
            texts.append(txt)
            alert.accept()
            handled = True
            time.sleep(0.3)
        except NoAlertPresentException:
            break
        except Exception:
            break
    return handled, texts


def get_payment_amount(driver: webdriver.Chrome, wait: WebDriverWait) -> Decimal:
    """
    读取“支払額”显示输入框的值：id=payment_formattedValue
    如出现 alert，先 accept 再读取
    """
    for _ in range(3):
        try:
            el = wait.until(EC.presence_of_element_located((By.ID, "payment_formattedValue")))
            val = el.get_attribute("value") or ""
            return to_decimal_money(val)
        except UnexpectedAlertPresentException:
            accept_any_alert(driver, max_rounds=5)
            time.sleep(0.3)

    el = wait.until(EC.presence_of_element_located((By.ID, "payment_formattedValue")))
    val = el.get_attribute("value") or ""
    return to_decimal_money(val)


def apply_invoice_by_autoenter(driver: webdriver.Chrome, wait: WebDriverWait, invoice_no: str) -> str:
    """
    在“アイテム選択(autoenter)”输入請求書番号并回车。
    回车后如有弹窗则自动 accept。
    返回：本次操作出现的弹窗文本（多弹窗用 " | " 拼接；无则空字符串）
    """
    el = wait.until(EC.element_to_be_clickable((By.ID, "autoenter")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.2)

    el.clear()
    el.send_keys(str(invoice_no))
    el.send_keys(Keys.ENTER)

    # 回车后优先处理可能弹出的 alert
    time.sleep(0.3)
    handled, texts = accept_any_alert(driver, max_rounds=5)
    if handled:
        return " | ".join([t.strip() for t in texts if str(t).strip()])
    return ""


def main():
    # 1) 读取Excel
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"未找到Excel：{EXCEL_PATH}")

    df = pd.read_excel(EXCEL_PATH, dtype={"請求書番号": str})
    required_cols = ["請求書番号", "請求書金額"]
    for c in required_cols:
        if c not in df.columns:
            raise ValueError(f"Excel缺少列：{c}（需要：{required_cols}）")

    invoice_pairs = []
    for _, row in df.iterrows():
        inv = (row.get("請求書番号") or "").strip()
        if inv == "" or inv.lower() == "nan":
            continue
        amt = to_decimal_money(row.get("請求書金額"))
        invoice_pairs.append((inv, amt))

    if not invoice_pairs:
        raise ValueError("Excel中没有有效的「請求書番号」数据。")

    # Excel 合计（用于最终总校验）
    excel_total = sum((amt for _, amt in invoice_pairs), Decimal("0.00")).quantize(Decimal("0.01"))

    # 日志文件
    log_path = os.path.join(LOG_DIR, f"apply_check_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

    driver = build_driver()
    wait = WebDriverWait(driver, 30)

    log_rows = []
    bad_invoices = []  # 记录金额有误的請求書番号

    try:
        # 1) 打开URL
        driver.get(URL)

        # 2) 手动登录
        wait_for_manual_login()

        # 确保关键字段已出现（未登录会等不到）
        wait.until(EC.presence_of_element_located((By.ID, "payment_formattedValue")))
        wait.until(EC.presence_of_element_located((By.ID, "autoenter")))

        # 3) 循环适用
        for invoice_no, expected_amt in invoice_pairs:
            # 4) 输入前读取支払額
            before_amt = get_payment_amount(driver, wait)

            # 5) 输入請求書番号并回车（并捕获弹窗文本）
            alert_text = apply_invoice_by_autoenter(driver, wait, invoice_no)

            # 6) 再次读取支払額，并校验差额
            #    勾选/刷新可能有延迟：等待金额变化（最多 8 秒）
            after_amt = None
            start = time.time()
            while True:
                try:
                    cur = get_payment_amount(driver, wait)
                except UnexpectedAlertPresentException:
                    accept_any_alert(driver, max_rounds=5)
                    cur = get_payment_amount(driver, wait)

                after_amt = cur
                if cur != before_amt:
                    break
                if time.time() - start > 8:
                    break
                time.sleep(0.3)

            diff = (after_amt - before_amt).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

            ok = (diff - expected_amt).copy_abs() <= TOLERANCE
            result = "金额无误" if ok else f"{invoice_no} 金额有误"

            if not ok:
                bad_invoices.append(invoice_no)

            log_rows.append({
                "timestamp": now_ts(),
                "請求書番号": invoice_no,
                "請求書金額(expected)": str(expected_amt),
                "支払額_before": str(before_amt),
                "支払額_after": str(after_amt),
                "差额(diff)": str(diff),
                "判定": result,
                "alert_text": alert_text,  # ★ 弹窗内容
            })

        # 7) 输出逐笔校验日志
        out_df = pd.DataFrame(log_rows)
        out_df.to_csv(log_path, index=False, encoding="utf-8-sig")
        print(f"\n完成。逐笔校验日志已输出：{log_path}")

        # ========= 8) 回车前：总校验 & 高亮提示 =========
        final_payment = get_payment_amount(driver, wait)
        total_ok = (final_payment - excel_total).copy_abs() <= TOLERANCE

        print("\n================= 回车前总校验 =================")
        print(f"Excel「請求書金額」合计 : {excel_total}")
        print(f"页面最终「支払額」       : {final_payment}")

        if total_ok:
            print("总校验结果：支払額 = Excel 合计 ✅")
        else:
            delta = (final_payment - excel_total).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            print(c_red(f"总校验结果：支払額 ≠ Excel 合计 ❌（差额={delta}）"))

        if bad_invoices:
            print(c_red(f"发现 {len(bad_invoices)} 条「金额有误」："))
            preview = bad_invoices[:50]
            line = "、".join(preview) + (" …" if len(bad_invoices) > 50 else "")
            print(c_yellow(line))
        else:
            print("逐笔校验结果：无「金额有误」记录 ✅")

        print("================================================\n")

        input("请确认页面无误（必要时可手动保存/截图）。按回车键结束程序并关闭浏览器：")

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
