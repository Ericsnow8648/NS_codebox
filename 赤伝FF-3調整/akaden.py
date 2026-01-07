# -*- coding: utf-8 -*-
"""
akaden.py

赤伝（クレジットメモ）を一括で編集する RPA スクリプト。

フォルダ構成（例）:
    C:/Users/Owner/Desktop/rpa_akaden/
        ├─ akaden.py
        ├─ akaden.xlsx   ... 入力用Excel（内部ID列 必須）
        └─ log.txt       ... エラーログ（無ければ自動作成）

Excel 仕様:
    - シート内に「内部ID」列があること
    - その他の列があっても無視されます

処理フロー:
    1. akaden.xlsx から内部IDリストを読み込む
    2. NetSuite に手動ログイン
    3. 各内部IDごとに:
        - 対象URLにアクセス
        - 「編集」ボタン押下
        - 編集画面ロード後に出るかもしれない alert を自動 OK
        - メモ欄に「FF-3処理済み」
        - 場所を「弁天倉庫」に変更（alert が出れば OK）
        - アイテムテーブル item_splits の各行について:
            - 在庫詳細アイコン（ダンボール）をクリックして行を展開
            - 展開行の inventorydetail_helper_popup をクリック
            - 在庫詳細ポップアップで
                ・保管棚 FF-3
                ・在庫ステータス 不良品
                ・数量を元数量にセット
                ・OK
        - 保存ボタン押下
        - 「保存されました」メッセージを待つ
    4. 失敗・例外は log.txt に書き出す
"""

import os
import time
import traceback

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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
EXCEL_FILE = os.path.join(BASE_DIR, "akaden.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "log.txt")

# ★ 赤伝の「表示」画面のURLベース
BASE_URL = (
    "https://6806569.app.netsuite.com/app/accounting/transactions/custcred.nl?id="
)


# =========================
# ログ関数
# =========================
def log_error(internal_id, reason=""):
    """log.txt にエラーを書き出す"""
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        ts = time.strftime("[%Y-%m-%d %H:%M:%S]")
        f.write(f"{ts} 内部ID={internal_id} {reason}\n")


def handle_possible_alert(driver, timeout=3, internal_id=None, context="", log=True):
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
        # alert が出なかったケース
        pass


# =========================
# 在庫詳細ポップアップ処理（iframe 版）
# =========================
def process_inventory_detail_popup(driver, internal_id, row_idx):
    """
    1行分の在庫詳細ポップアップで
      - 保管棚: FF-3
      - ステータス: 不良品
      - 数量: 元数量
      - OK で閉じる
    """

    wait = WebDriverWait(driver, 5)

    # ----- iframe に入る -----
    try:
        wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.NAME, "childdrecord_frame")
            )
        )
    except TimeoutException:
        wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.ID, "childdrecord_frame")
            )
        )

    try:
        # 元数量を取得（無ければ "1"）
        assign_qty = "1"
        try:
            q_span = wait.until(
                EC.presence_of_element_located((By.ID, "quantity_val"))
            )
            q_text = (q_span.text or "").strip()
            if q_text:
                assign_qty = q_text
        except TimeoutException:
            pass

        # ---- 保管棚 FF-3 ----
        bin_input = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "inventoryassignment_binnumber_display")
            )
        )
        time.sleep(0.2)
        try:
            bin_input.click()
        except Exception:
            driver.execute_script("arguments[0].click();", bin_input)
        time.sleep(0.2)

        bin_input.send_keys(Keys.CONTROL, "a")
        bin_input.send_keys(Keys.DELETE)
        time.sleep(0.1)
        bin_input.send_keys("FF-3")
        time.sleep(0.3)

        # TAB でステータス欄へ移動
        bin_input.send_keys(Keys.TAB)
        time.sleep(0.5)

        # ---- ステータス 不良品 ----
        status_input = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[id^='inpt_inventorystatus']")
            )
        )
        time.sleep(0.3)

        try:
            status_input.click()
        except Exception:
            driver.execute_script("arguments[0].click();", status_input)
        time.sleep(0.4)

        # 「通常在庫」から ARROW_UP で不良品まで上がっていく
        for _ in range(10):
            current_val = status_input.get_attribute("value") or ""
            if "不良品" in current_val:
                break
            status_input.send_keys(Keys.ARROW_UP)
            time.sleep(0.25)

        status_input.send_keys(Keys.ENTER)  # 不良品で確定
        time.sleep(0.4)

        # ★ ここが重要: 数量欄へ移動する TAB は status_input に送る
        status_input.send_keys(Keys.TAB)
        time.sleep(0.5)

        # ---- 数量入力 ----
        try:
            qty_input = wait.until(
                EC.element_to_be_clickable(
                    (By.ID, "quantity_formattedValue")
                )
            )
            time.sleep(0.1)
            try:
                qty_input.click()
            except Exception:
                driver.execute_script("arguments[0].click();", qty_input)
            time.sleep(0.1)

            qty_input.send_keys(Keys.CONTROL, "a")
            qty_input.send_keys(Keys.DELETE)
            qty_input.send_keys(assign_qty)
            time.sleep(0.2)
            qty_input.send_keys(Keys.TAB)  # blur させて確定
            time.sleep(0.3)
        except TimeoutException:
            # 数量欄が無いケースはそのまま進む
            log_error(internal_id, "数量入力欄(quantity_formattedValue)が見つからずスキップ")

        # ---- 行内 OK ボタン ----
        try:
            ok_line = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.ID, "inventoryassignment_addedit")
                )
            )
            try:
                ok_line.click()
            except Exception:
                driver.execute_script("arguments[0].click();", ok_line)
            time.sleep(0.5)
        except TimeoutException:
            pass

        # ---- ポップアップ全体の OK ----
        ok_popup = wait.until(
            EC.element_to_be_clickable((By.ID, "secondaryok"))
        )
        try:
            ok_popup.click()
        except Exception:
            driver.execute_script("arguments[0].click();", ok_popup)
        time.sleep(0.5)

    except Exception as e:
        log_error(internal_id, f"在庫詳細ポップアップ(row={row_idx})で例外: {e}")
        raise
    finally:
        # iframe から親に戻る
        try:
            driver.switch_to.default_content()
        except Exception:
            pass



# =========================
# メイン処理
# =========================
def main():
    # ---------- Excel 読み込み ----------
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Excel ファイルが見つかりません: {EXCEL_FILE}")

    df = pd.read_excel(EXCEL_FILE)
    if "内部ID" not in df.columns:
        raise ValueError("Excel に '内部ID' 列が必要です")

    df = df.dropna(subset=["内部ID"])
    df["内部ID"] = df["内部ID"].astype(str).str.strip()
    records = sorted(set(df["内部ID"].tolist()))

    if not records:
        print("処理対象の内部IDがありません。")
        return

    print(f"対象件数: {len(records)} 件")

    # ---------- Chrome 起動 ----------
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )
    driver.maximize_window()

    # NetSuite ログイン
    driver.get("https://6806569.app.netsuite.com")
    input("🔐 NetSuite にログイン完了後、Enter を押してください...")

    wait = WebDriverWait(driver, 20)

    # ---------- 内部IDごとの処理 ----------
    for internal_id in records:
        print(f"\n===== 開始: 内部ID={internal_id} =====")
        try:
            url = BASE_URL + str(internal_id)
            driver.get(url)

            # 「編集」ボタン待ち
            edit_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "edit"))
            )
            time.sleep(0.5)
            try:
                edit_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", edit_btn)

            # 編集画面のロード待ち（保存ボタンが出るまで）
            wait.until(
                EC.presence_of_element_located(
                    (By.ID, "btn_secondarymultibutton_submitter")
                )
            )

            # 編集直後の alert を処理
            handle_possible_alert(
                driver,
                timeout=5,
                internal_id=internal_id,
                context="after_edit",
                log=False,
            )

            # ---------- メモ: FF-3処理済み ----------
            try:
                memo_input = wait.until(
                    EC.presence_of_element_located((By.ID, "memo"))
                )
                time.sleep(0.2)
                memo_input.click()
                time.sleep(0.1)
                memo_input.send_keys(Keys.CONTROL, "a")
                memo_input.send_keys(Keys.DELETE)
                memo_input.send_keys("FF-3処理済み")
                time.sleep(0.1)
                memo_input.send_keys(Keys.TAB)
                time.sleep(0.2)
            except Exception as e:
                log_error(internal_id, f"メモ入力で例外: {e}")

            # ---------- 場所: 弁天倉庫 ----------
            try:
                loc_input = wait.until(
                    EC.element_to_be_clickable((By.ID, "location_display"))
                )
                time.sleep(0.2)
                loc_input.click()
                time.sleep(0.1)
                loc_input.send_keys(Keys.CONTROL, "a")
                loc_input.send_keys(Keys.DELETE)
                time.sleep(0.1)
                loc_input.send_keys("弁天倉庫")
                time.sleep(0.6)
                loc_input.send_keys(Keys.ARROW_DOWN)
                time.sleep(0.2)
                loc_input.send_keys(Keys.ENTER)
                time.sleep(0.8)

                # 場所変更の確認 alert
                handle_possible_alert(
                    driver,
                    timeout=5,
                    internal_id=internal_id,
                    context="after_location_change",
                    log=False,
                )

            except Exception as e:
                log_error(internal_id, f"場所(弁天倉庫)の設定で例外: {e}")

            # ---------- アイテム行ループ（item_splits） ----------
            try:
                table = wait.until(
                    EC.presence_of_element_located((By.ID, "item_splits"))
                )

                row_idx = 1
                while True:
                    # 行を探す
                    try:
                        row = table.find_element(By.ID, f"item_row_{row_idx}")
                    except NoSuchElementException:
                        break  # これ以上行がない

                    try:
                        # 在庫詳細アイコン（灰色ダンボール）をクリックして編集行を出す
                        icon_span = row.find_element(
                            By.CSS_SELECTOR,
                            "span.uir-helper-icon.smalltextul.field_widget.i_inventorydetailneeded"
                        )
                    except NoSuchElementException:
                        print(f"  [行{row_idx}] 在庫詳細アイコンなし → スキップ")
                        row_idx += 1
                        continue

                    print(f"  [行{row_idx}] 在庫詳細アイコンクリック")
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'});", icon_span
                    )
                    time.sleep(0.2)
                    try:
                        icon_span.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", icon_span)
                    time.sleep(0.8)

                    # 展開された行の中の inventorydetail_helper_popup（青いダンボール）をクリック
                    try:
                        inv_link = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable(
                                (By.ID, "inventorydetail_helper_popup")
                            )
                        )
                        driver.execute_script(
                            "arguments[0].scrollIntoView({block:'center'});",
                            inv_link,
                        )
                        time.sleep(0.3)
                        try:
                            inv_link.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", inv_link)
                        time.sleep(1.0)

                        # ポップアップ内処理
                        process_inventory_detail_popup(driver, internal_id, row_idx)
                        print(f"  [行{row_idx}] 在庫詳細処理完了")

                    except TimeoutException as e_popup:
                        log_error(
                            internal_id,
                            f"行{row_idx} 在庫詳細ポップアップ起動でタイムアウト: {e_popup}",
                        )
                        print(f"  ❌ 行{row_idx} 在庫詳細ポップアップ起動失敗")
                    except Exception as e_row:
                        log_error(
                            internal_id,
                            f"行{row_idx} 在庫詳細処理で例外: {e_row}",
                        )
                        print(f"  ❌ 行{row_idx} 在庫詳細処理でエラー: {e_row}")

                    finally:
                        row_idx += 1

            except Exception as e:
                log_error(internal_id, f"item_splits テーブル処理で例外: {e}")

            # ---------- 保存 ----------
            try:
                save_btn = wait.until(
                    EC.element_to_be_clickable(
                        (By.ID, "btn_secondarymultibutton_submitter")
                    )
                )
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});", save_btn
                )
                time.sleep(0.2)
                try:
                    save_btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", save_btn)

                # 保存後の warning alert を処理
                handle_possible_alert(
                    driver,
                    timeout=10,
                    internal_id=internal_id,
                    context="after_save_click",
                    log=True,
                )

                # 「保存されました」メッセージ待機
                try:
                    WebDriverWait(driver, 20).until(
                        EC.text_to_be_present_in_element(
                            (By.CSS_SELECTOR, "div.content div.descr"),
                            "保存されました",
                        )
                    )
                    print(f"✅ 完了: 内部ID={internal_id}")
                except TimeoutException:
                    log_error(internal_id, "保存メッセージ確認できず（タイムアウト）")
                    print(f"⚠️ 保存メッセージ確認できず: 内部ID={internal_id}")

            except Exception as e:
                log_error(internal_id, f"保存ボタンクリックで例外: {e}")
                print(f"❌ 保存失敗: 内部ID={internal_id} -> {e}")
                continue

        except UnexpectedAlertPresentException:
            try:
                alert = driver.switch_to.alert
                msg = alert.text
                alert.accept()
            except Exception:
                msg = "alert-handling-failed"
            log_error(internal_id, f"UnexpectedAlert: {msg}")
            print(f"🚨 Unexpected alert: 内部ID={internal_id} -> {msg}")
            continue

        except Exception as e:
            log_error(internal_id, f"例外: {e}\n{traceback.format_exc()}")
            print(f"❌ エラー: 内部ID={internal_id} -> {e}")
            continue

    # ---------- 終了処理 ----------
    driver.quit()
    print("\n🏁 全ての処理が完了しました。")


if __name__ == "__main__":
    main()
