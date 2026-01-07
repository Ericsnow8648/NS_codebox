# -*- coding: utf-8 -*-
"""
seikyuu_auto_department.py

機能:
    Excel「注文書.xlsx」の「内部ID」「日期」「顾客」列を読み込み、
    Sales Order(注文書) を開いて「請求」ボタンを押し、
    新規請求書画面で:

      1. 日付(trandate) を Excel の「日期」に変更
      2. 顧客コードに応じて「部門」を自動選択
         - C000222 → EC (BtoC）
         - C000142 → 営業(BtoB）
      3. 保存

Excel 必須列:
    ・「内部ID」
    ・「日期」
    ・「顾客」
"""

import os
import time
import traceback
import datetime as dt

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    UnexpectedAlertPresentException,
)
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# ==========================
# 設定
# ==========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 只从この Excel 读取数据
EXCEL_FILE = os.path.join(BASE_DIR, "注文書.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "log_seikyuu_department.txt")

# Sales Order (注文書) 基础 URL
BASE_URL_SALESORD = (
    "https://6806569.app.netsuite.com/app/accounting/transactions/salesord.nl?id="
)

# 登录用 URL（任意能打开的订单页面）
LOGIN_START_URL = (
    "https://6806569.app.netsuite.com/app/accounting/transactions/salesord.nl?"
    "id=6875021&whence="
)

# 顧客 → 部門的映射（这里只用来提示/日志，不直接用于操作下拉）
CUSTOMER_TO_DEPARTMENT = {
    "C000222": "EC (BtoC）",
    "C000142": "営業(BtoB）",
}


# ==========================
# 共通工具函数
# ==========================
def log_error(internal_id, reason=""):
    """写入错误日志"""
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        ts = time.strftime("[%Y-%m-%d %H:%M:%S]")
        f.write(f"{ts} 内部ID={internal_id} {reason}\n")


def handle_possible_alert(driver, timeout=5, internal_id=None, context="", log=True):
    """
    一定时间内如果出现 alert，则接受；否则什么也不做
    """
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        txt = alert.text
        alert.accept()
        msg = f"[Alert {context}] {txt}"
        print("⚠️", msg)
        if log and internal_id:
            log_error(internal_id, msg)
        time.sleep(0.4)
    except TimeoutException:
        pass


def format_date_for_ns(value):
    """Excel日期 → NetSuite 日期格式 yyyy/mm/dd"""
    if pd.isna(value):
        return ""

    if isinstance(value, (pd.Timestamp, dt.datetime, dt.date)):
        return value.strftime("%Y/%m/%d")

    return str(value).strip()


# ==========================
# 注文書 → 請求書
# ==========================
def click_bill_button(driver, wait, internal_id):
    """在 Sales Order 画面点击「請求」按钮 (id='billremaining')"""
    try:
        btn = wait.until(
            EC.element_to_be_clickable((By.ID, "billremaining"))
        )
    except TimeoutException as e:
        msg = "注文書画面找不到『請求』按钮 (id='billremaining')"
        log_error(internal_id, msg)
        raise TimeoutException(msg) from e

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", btn
    )
    time.sleep(0.3)

    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)

    time.sleep(1.2)


# ==========================
# 請求書页面操作
# ==========================
def set_trandate(driver, internal_id, date_str):
    """设置請求書的日付(trandate)"""
    if not date_str:
        log_error(internal_id, "Excel 日期为空，跳过日付设置")
        return

    try:
        inp = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "trandate"))
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", inp
        )
        time.sleep(0.3)

        inp.click()
        time.sleep(0.2)
        inp.send_keys(Keys.CONTROL, "a")
        inp.send_keys(Keys.DELETE)
        time.sleep(0.1)
        inp.send_keys(date_str)
        time.sleep(0.2)
        inp.send_keys(Keys.ENTER)
        time.sleep(0.5)

        print(f"✅ 日付设置完成: {date_str}")

    except Exception as e:
        msg = f"日付输入失败: {e}"
        print(f"❌ {msg}")
        log_error(internal_id, msg)


def set_department_by_customer(driver, internal_id, customer_code):
    """
    根据顾客选择部门（通过下拉列表逐个移动选择）:

    部門选项顺序（来自你给的 data-options）:
        0: ""                  （空）
        1: "EC (BtoC）"
        2: "アウトドアプロジェクト"
        3: "営業(BtoB）"
        4: "東京オフィス"
        5: "管理部"
        6: "輸出事業部"
        ...

    对应规则:
        顾客 C000222 → index 1 → 从顶部往下移动 1 次
        顾客 C000142 → index 3 → 从顶部往下移动 3 次
    """

    customer_code = (customer_code or "").strip()

    # 顾客 → 需要按 DOWN 的次数（从 HOME 后的第 0 项 = 空开始）
    customer_to_steps = {
        "C000222": 1,  # EC (BtoC）
        "C000142": 3,  # 営業(BtoB）
    }

    if customer_code not in customer_to_steps:
        msg = f"顾客 {customer_code} 无对应部門 index 映射，跳过部門设置"
        print("⚠️", msg)
        log_error(internal_id, msg)
        return

    steps = customer_to_steps[customer_code]
    dept_label = CUSTOMER_TO_DEPARTMENT.get(customer_code, f"(steps={steps})")
    print(f"➡️ 设置部門：顾客={customer_code}, 目标='{dept_label}', 从顶部往下移动 {steps} 次")

    try:
        # 锁定可输入的下拉 input[name='inpt_department']
        dept_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//input[@name='inpt_department' and contains(@id,'inpt_department')]",
                )
            )
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", dept_input
        )
        time.sleep(0.3)

        # 点击激活下拉
        dept_input.click()
        time.sleep(0.3)

        actions = ActionChains(driver)

        # 清空当前文本
        actions.send_keys(Keys.CONTROL, "a")
        actions.send_keys(Keys.DELETE)
        actions.pause(0.2)

        # HOME：移动到列表最上面的空白项(索引0)
        actions.send_keys(Keys.HOME)
        actions.pause(0.2)

        # 向下移动 steps 次
        for _ in range(steps):
            actions.send_keys(Keys.DOWN)
            actions.pause(0.15)

        # ENTER 选中
        actions.send_keys(Keys.ENTER)
        actions.perform()

        time.sleep(0.6)

        # 可选：读取隐藏字段，确认已设置
        try:
            hidden_val = driver.execute_script(
                "var el = document.querySelector(\"input[name='department']\");"
                "return el ? el.value : null;"
            )
            print(f"   🔎 hidden department value = {hidden_val}")
        except Exception:
            pass

        print(f"✅ 部門选择完成（下移 {steps} 次）")

    except Exception as e:
        msg = f"设置部门失败（HOME+DOWN 方式）: {e}"
        print(f"❌ {msg}")
        log_error(internal_id, msg)


def save_invoice(driver, wait, internal_id):
    """保存請求書"""
    try:
        btn = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "btn_secondarymultibutton_submitter")
            )
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", btn
        )
        time.sleep(0.3)

        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)

        handle_possible_alert(driver, timeout=10, internal_id=internal_id)

        WebDriverWait(driver, 25).until(
            EC.text_to_be_present_in_element(
                (By.CSS_SELECTOR, "div.content div.descr"),
                "保存されました",
            )
        )
        print(f"✅ 請求書保存完成：内部ID={internal_id}")

    except Exception as e:
        msg = f"保存失败: {e}"
        print(f"❌ {msg}")
        log_error(internal_id, msg)
        # 不 raise，避免中断后面的记录


# ==========================
# 主流程
# ==========================
def main():
    # ---------- 读取 Excel ----------
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"找不到 Excel 文件: {EXCEL_FILE}")

    df = pd.read_excel(
        EXCEL_FILE,
        dtype={"内部ID": str, "顾客": str},
        keep_default_na=False,
    )

    required = ["内部ID", "日期", "顾客"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Excel 缺少必要列：{col}")

    df["内部ID"] = df["内部ID"].str.strip()
    df = df[df["内部ID"] != ""]

    df["日期文字列"] = df["日期"].apply(format_date_for_ns)

    records = df.to_dict("records")
    if not records:
        print("Excel 中没有需要处理的内部ID。")
        return

    print(f"目标件数：{len(records)} 件")

    # ---------- 启动浏览器 ----------
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )
    driver.maximize_window()
    wait = WebDriverWait(driver, 25)

    # ---------- 登录 ----------
    driver.get(LOGIN_START_URL)
    input("🔐 请在浏览器中完成 NetSuite 登录，然后按 Enter 继续...")

    # ---------- 按行处理 ----------
    for row in records:
        internal_id = row["内部ID"]
        date_str = row["日期文字列"]
        customer = row["顾客"]

        print(
            f"\n===== 开始处理：内部ID={internal_id} 顾客={customer} 日期={date_str} ====="
        )

        try:
            # 打开 Sales Order
            url = BASE_URL_SALESORD + internal_id
            driver.get(url)

            handle_possible_alert(
                driver,
                timeout=3,
                internal_id=internal_id,
                context="open_salesorder",
                log=False,
            )

            # 点击「請求」按钮
            click_bill_button(driver, wait, internal_id)

            handle_possible_alert(
                driver,
                timeout=5,
                internal_id=internal_id,
                context="after_bill_click",
                log=False,
            )

            # 设置日付
            set_trandate(driver, internal_id, date_str)

            # 设置部門
            set_department_by_customer(driver, internal_id, customer)

            # 保存請求書
            save_invoice(driver, wait, internal_id)

        except UnexpectedAlertPresentException:
            try:
                alert = driver.switch_to.alert
                msg = alert.text
                alert.accept()
            except Exception:
                msg = "alert-handling-failed"

            log_error(internal_id, f"Unexpected alert: {msg}")
            print(f"⚠️ Unexpected alert: {msg}")
            continue

        except Exception as e:
            msg = f"例外: {e}\n{traceback.format_exc()}"
            print(f"❌ 发生错误：{e}")
            log_error(internal_id, msg)
            continue

    driver.quit()
    print("\n🏁 全部处理完毕，如有错误请查看 log_seikyuu_department.txt")


if __name__ == "__main__":
    main()
