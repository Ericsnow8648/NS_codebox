# -*- coding: utf-8 -*-
"""
henpin_auto_akaden.py

返品（Return Authorization）から「払戻」を押して
赤伝（クレジットメモ）画面で:

1. 日付を Excel の値に変更（該当 返品内部ID の行の「日付」列）
2. 「適用」タブの「アイテム選択」に 請求書番号 を入力して Enter
3. 保存

を自動実行する RPA スクリプト。

Excel 仕様:
    - 少なくとも以下の列があること:
        ・「返品内部ID」   … Return Authorization の internalid
        ・「日付」         … 赤伝に設定したい日付
        ・「請求書番号」   … 「アイテム選択」に入力する請求書番号
        ・「金額」         … 0 なら「適用」スキップ、それ以外は従来通り
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
    NoSuchElementException,
    UnexpectedAlertPresentException,
)
from webdriver_manager.chrome import ChromeDriverManager

# =========================
# 設定
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "henpin.xlsx")          # 入力 Excel
LOG_FILE = os.path.join(BASE_DIR, "log_henpin_akaden.txt")  # ログ

# 返品（Return Authorization）の表示 URL ベース
BASE_URL_RTNAUTH = (
    "https://6806569.app.netsuite.com/app/accounting/transactions/rtnauth.nl?id="
)

# =========================
# ログ関数
# =========================
def log_error(internal_id, reason=""):
    """log_henpin_akaden.txt にエラーを書き出す"""
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        ts = time.strftime("[%Y-%m-%d %H:%M:%S]")
        f.write(f"{ts} 返品内部ID={internal_id} {reason}\n")


def handle_possible_alert(driver, timeout=5, internal_id=None, context="", log=True):
    """
    一定時間内に alert が出ていれば OK を押して閉じる。
    出なければ何もしない。
    """
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        text = alert.text
        alert.accept()
        msg = f"Alert[{context}] -> {text}"
        print("⚠️", msg)
        if log and internal_id is not None:
            log_error(internal_id, msg)
        time.sleep(0.5)
    except TimeoutException:
        pass


# =========================
# 返品画面 → 赤伝画面へ
# =========================
def click_refund_button(driver, wait, internal_id):
    """
    返品(表示)画面で「払戻」ボタン（input id='refund'）をクリックする。
    """

    try:
        refund_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "refund"))
        )
    except TimeoutException as e:
        msg = "返品画面で id='refund'（払戻）ボタンが見つかりません。"
        log_error(internal_id, msg)
        raise TimeoutException(msg) from e

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", refund_btn
    )
    time.sleep(0.3)
    try:
        refund_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", refund_btn)
    time.sleep(1.0)  # 赤伝画面への遷移待ち


# =========================
# 日付文字列整形
# =========================
def format_date_for_ns(value):
    """
    Excel の「日付」セルから NetSuite に入力するための文字列に整形。
    NetSuite の UI では通常 'yyyy/mm/dd' 形式が安定。
    """
    if pd.isna(value):
        return ""

    # Pandas Timestamp / datetime.date / datetime.datetime の場合
    if isinstance(value, (pd.Timestamp, dt.datetime, dt.date)):
        return value.strftime("%Y/%m/%d")

    # 文字列の場合はそのまま返す（ユーザーが Excel 側でフォーマットを揃える前提）
    s = str(value).strip()
    return s


# =========================
# 赤伝画面の操作（日付 + アイテム選択 + 保存）
# =========================
def process_credit_memo(driver, wait, internal_id, date_str, invoice_no, need_apply=True):
    """
    赤伝画面でやること：
      1. 日付(trandate) を date_str に変更し Enter
      2. （金額 != 0 の場合のみ）
         「適用」タブを開き、「アイテム選択(autoenter)」に請求書番号を入力して Enter
      3. 保存ボタン押下
    """

    # 画面ロード確認（保存ボタンが出るまで待つ）
    save_btn = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "btn_secondarymultibutton_submitter")
        )
    )

    # ========= 1) 日付の入力 =========
    if date_str:
        try:
            date_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "trandate"))
            )

            # ★ 日付欄を画面中央にスクロールして見えるようにする
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", date_input
            )
            time.sleep(0.3)

            try:
                date_input.click()
            except Exception:
                driver.execute_script("arguments[0].click();", date_input)
            time.sleep(0.2)

            date_input.send_keys(Keys.CONTROL, "a")
            date_input.send_keys(Keys.DELETE)
            time.sleep(0.1)
            date_input.send_keys(date_str)
            time.sleep(0.2)
            date_input.send_keys(Keys.ENTER)  # 入力確定
            time.sleep(0.5)

        except Exception as e:
            log_error(internal_id, f"赤伝画面の日付(trandate)入力で例外: {e}")

    else:
        log_error(internal_id, "Excel の『日付』が空のため、日付変更をスキップ")

    # ========= 2) 「適用」タブ → アイテム選択(autoenter) に請求書番号を入力 =========
    # ★ 金額が 0 の場合は「適用」タブの操作を丸ごとスキップ
    if need_apply:
        if invoice_no:
            try:
                # 適用タブをクリック（ID は applytxt が多い）
                try:
                    apply_tab = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "applytxt"))
                    )

                    # ★ 適用タブも中央付近にスクロールしておくと見やすい
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'});", apply_tab
                    )
                    time.sleep(0.3)

                    try:
                        apply_tab.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", apply_tab)
                    time.sleep(0.5)
                except TimeoutException:
                    # タブクリックに失敗したらそのまま続行（既に適用タブが開いている可能性）
                    pass

                # アイテム選択の入力欄（id="autoenter"）
                auto_input = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "autoenter"))
                )

                # ★ アイテム選択欄も画面中央にスクロール
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});", auto_input
                )
                time.sleep(0.3)

                try:
                    auto_input.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", auto_input)
                time.sleep(0.2)

                auto_input.send_keys(Keys.CONTROL, "a")
                auto_input.send_keys(Keys.DELETE)
                time.sleep(0.1)
                auto_input.send_keys(str(invoice_no))
                time.sleep(0.2)
                auto_input.send_keys(Keys.ENTER)  # アイテム選択確定
                time.sleep(1.0)

            except Exception as e:
                log_error(internal_id, f"赤伝画面のアイテム選択(autoenter)入力で例外: {e}")
        else:
            log_error(internal_id, "Excel の『請求書番号』が空のため、アイテム選択をスキップ")
    else:
        # 金額 0 の場合は適用タブの操作を行わない
        print(f"ℹ️ 金額=0 のため、適用タブのアイテム選択をスキップします: 返品内部ID={internal_id}")

    # ========= 3) 保存 =========
    try:
        # ★ 最後に保存ボタンも中央に持ってきておくと、保存の瞬間も目視しやすい
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", save_btn
        )
        time.sleep(0.3)
        try:
            save_btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", save_btn)

        # 保存後の alert 処理
        handle_possible_alert(
            driver,
            timeout=10,
            internal_id=internal_id,
            context="credit_memo_save_click",
            log=True,
        )

        # 「保存されました」メッセージ待機
        WebDriverWait(driver, 20).until(
            EC.text_to_be_present_in_element(
                (By.CSS_SELECTOR, "div.content div.descr"),
                "保存されました",
            )
        )
        print(f"✅ 赤伝保存完了: 返品内部ID={internal_id}")

    except TimeoutException:
        msg = "赤伝保存後のメッセージ『保存されました』が確認できませんでした（タイムアウト）。"
        log_error(internal_id, msg)
        print(f"⚠️ {msg} 返品内部ID={internal_id}")
    except Exception as e:
        log_error(internal_id, f"赤伝保存処理で例外: {e}")
        print(f"❌ 赤伝保存失敗: 返品内部ID={internal_id} -> {e}")
        raise



# =========================
# メイン処理
# =========================
def main():
    # ---------- Excel 読み込み ----------
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Excel ファイルが見つかりません: {EXCEL_FILE}")

    # ★ 在读取时就指定「返品内部ID」「請求書番号」为字符串
    df = pd.read_excel(
        EXCEL_FILE,
        dtype={
            "返品内部ID": str,
            "請求書番号": str,
        },
        # 可选：避免空单元格变成 NaN 字符串
        keep_default_na=False,
    )

    # ★ 必須列に「金額」も追加
    required_cols = ["返品内部ID", "日付", "請求書番号", "金額"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Excel に '{col}' 列が必要です")

    # 过滤掉 返品内部ID 为空的行（先 strip 再判断）
    df["返品内部ID"] = df["返品内部ID"].astype(str).str.strip()
    df["請求書番号"] = df["請求書番号"].astype(str).str.strip()

    df = df[df["返品内部ID"] != ""]

    # ★ 金額列 → 数値化（空白或非法值视为 0）
    df["金額"] = pd.to_numeric(df["金額"], errors="coerce").fillna(0)

    # 日付列 → 文字列列（用你之前的 format_date_for_ns）
    df["日付文字列"] = df["日付"].apply(format_date_for_ns)

    records = df.to_dict("records")

    if not records:
        print("処理対象の返品内部IDがありません。")
        return

    print(f"対象返品件数: {len(records)} 件")

    # ---------- Chrome 起動 ----------
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )
    driver.maximize_window()
    wait = WebDriverWait(driver, 20)

    # NetSuite ログイン
    driver.get("https://6806569.app.netsuite.com")
    input("🔐 NetSuite にログイン完了後、Enter を押してください...")

    # ---------- 返品ごとの処理 ----------
    for row in records:
        return_id = row["返品内部ID"]
        date_str = row["日付文字列"]
        invoice_no = row["請求書番号"]
        amount = row.get("金額", 0)

        # ★ 金額是否为 0 决定是否需要「適用」操作
        try:
            amount_val = float(amount)
        except Exception:
            amount_val = 0.0

        need_apply = (amount_val != 0.0)

        print(
            f"\n===== 開始: 返品内部ID={return_id} 日付={date_str} "
            f"請求書番号={invoice_no} 金額={amount_val} need_apply={need_apply} ====="
        )

        try:
            # 返品（Return Authorization）の「表示」画面へ
            url = BASE_URL_RTNAUTH + str(return_id)
            driver.get(url)

            # ★ 打开返品画面后，先把可能出现的「締め請求書を使用」等信息弹窗关掉
            handle_possible_alert(
                driver,
                timeout=3,
                internal_id=return_id,
                context="open_return",
                log=False,  # 纯信息不记日志
            )

            # main_form 等のロード待ち（軽く）
            try:
                wait.until(
                    EC.presence_of_element_located((By.ID, "main_form"))
                )
            except TimeoutException:
                pass

            # 1) 返品画面の「払戻」ボタンをクリック → 赤伝画面へ
            click_refund_button(driver, wait, return_id)

            handle_possible_alert(
                driver,
                timeout=3,
                internal_id=return_id,
                context="after_refund_click",
                log=False,
            )

            # 2) 赤伝画面で 日付 + （必要なら）アイテム選択 + 保存
            process_credit_memo(
                driver,
                wait,
                return_id,
                date_str,
                invoice_no,
                need_apply=need_apply,
            )

        except UnexpectedAlertPresentException:
            # 有意外的 alert，就先把它关掉
            try:
                alert = driver.switch_to.alert
                msg = alert.text
                alert.accept()
            except Exception:
                msg = "alert-handling-failed"

            # 如果是「締め請求書を使用」相关的提醒，就当成信息提示，忽略
            if "締め請求書を使用" in msg:
                print(f"ℹ️ 締め請求書に関する情報アラートを無視して続行: 返品内部ID={return_id}")
                # 简单起见仍然跳到下一条
                continue

            # 其他未知的 alert 仍然当成错误处理
            log_error(return_id, f"UnexpectedAlert: {msg}")
            print(f"🚨 Unexpected alert: 返品内部ID={return_id} -> {msg}")
            continue

        except Exception as e:
            log_error(return_id, f"例外: {e}\n{traceback.format_exc()}")
            print(f"❌ エラー: 返品内部ID={return_id} -> {e}")
            continue

    # ---------- 終了処理 ----------
    driver.quit()
    print("\n🏁 全ての返品に対する赤伝処理が完了しました。（エラーは log_henpin_akaden.txt を確認）")


if __name__ == "__main__":
    main()
