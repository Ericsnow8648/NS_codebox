import os
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ================== 固定路径（按你给的地址） ==================
BASE_DIR = r"C:\Users\mitsu\OneDrive\デスクトップ\订单后续中间表 删除"
EXCEL_PATH = os.path.join(BASE_DIR, "delete_list.xlsx")
LOG_DIR = os.path.join(BASE_DIR, "logs")
# ============================================================

# ================== NetSuite 配置 ==================
COMPANY_BASE = "https://6806569.app.netsuite.com"
RECTYPE = 396
ID_COLUMN = "内部ID"
WAIT_SEC = 25
# ================================================


def build_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # 不复用 profile：让你手动登录
    driver = webdriver.Chrome(options=options)
    return driver


def wait_dom_ready(driver, timeout=WAIT_SEC):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def first_present(driver, xpaths, timeout=WAIT_SEC):
    end = time.time() + timeout
    last_err = None
    while time.time() < end:
        for xp in xpaths:
            try:
                el = driver.find_element(By.XPATH, xp)
                if el.is_displayed():
                    return el, xp
            except Exception as e:
                last_err = e
        time.sleep(0.2)
    raise TimeoutException(f"None present: {xpaths} ; last={last_err}")


def safe_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.15)
    el.click()


def try_accept_alert(driver, timeout=8):
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
        return True
    except TimeoutException:
        return False


def click_confirm_if_inline_modal(driver):
    confirm_xpaths = [
        "//button[normalize-space()='确定']",
        "//button[normalize-space()='OK']",
        "//button[normalize-space()='はい']",
        "//a[normalize-space()='确定']",
        "//a[normalize-space()='OK']",
        "//a[normalize-space()='はい']",
    ]
    try:
        el, _ = first_present(driver, confirm_xpaths, timeout=6)
        safe_click(driver, el)
        return True
    except Exception:
        return False


def open_record_edit(driver, internal_id: int):
    # 直接拼接编辑URL（&e=T）
    url = f"{COMPANY_BASE}/app/common/custom/custrecordentry.nl?rectype={RECTYPE}&id={internal_id}&e=T"
    driver.get(url)
    wait_dom_ready(driver)


def click_action_delete_on_edit_page(driver):
    # 1) アクション
    action_xpaths = [
        "//button[normalize-space()='アクション']",
        "//a[normalize-space()='アクション']",
        "//*[normalize-space()='アクション']/ancestor::button[1]",
        "//*[normalize-space()='アクション']/ancestor::a[1]",
    ]
    action_el, _ = first_present(driver, action_xpaths, timeout=WAIT_SEC)
    safe_click(driver, action_el)
    time.sleep(0.25)

    # 2) 削除
    delete_xpaths = [
        "//a[normalize-space()='削除']",
        "//span[normalize-space()='削除']/ancestor::a[1]",
        "//*[@role='menuitem' and normalize-space()='削除']",
        "//li[.//*[normalize-space()='削除']]",
        "//*[contains(@class,'dropdown') or contains(@class,'menu')]//*[normalize-space()='削除']",
    ]
    del_el, _ = first_present(driver, delete_xpaths, timeout=WAIT_SEC)
    safe_click(driver, del_el)


def wait_back_to_list(driver):
    def ok(d):
        t = (d.title or "")
        return ("リスト" in t) or ("List" in t) or ("订单后续中间表" in t)
    WebDriverWait(driver, WAIT_SEC).until(ok)


def read_ids_from_excel():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"找不到Excel：{EXCEL_PATH}")

    df = pd.read_excel(EXCEL_PATH)
    if ID_COLUMN not in df.columns:
        raise ValueError(f"Excel里找不到列：{ID_COLUMN}")

    ids = (
        df[ID_COLUMN]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s != ""]
        .astype(int)
        .tolist()
    )
    return ids


def main():
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"delete_log_{int(time.time())}.csv")

    ids = read_ids_from_excel()
    print(f"Excel路径: {EXCEL_PATH}")
    print(f"待删除数量: {len(ids)}")

    driver = build_driver()

    # 1) 打开 NetSuite 首页/登录页，手动登录
    driver.get(COMPANY_BASE)
    wait_dom_ready(driver)

    print("\n请在打开的浏览器中【手动登录 NetSuite】并确保能正常打开记录页面。")
    input("登录完成后，回到这里按【回车】开始批量删除...")

    results = []
    try:
        for internal_id in ids:
            row = {
                "内部ID": internal_id,
                "status": "",
                "message": "",
                "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
            }
            try:
                # 2) 直接进入编辑页
                open_record_edit(driver, internal_id)

                # 3) アクション -> 削除
                click_action_delete_on_edit_page(driver)

                # 4) 弹窗确定
                accepted = try_accept_alert(driver, timeout=8)
                if not accepted:
                    click_confirm_if_inline_modal(driver)

                # 5) 回到列表页
                wait_back_to_list(driver)

                row["status"] = "OK"
                row["message"] = "Deleted"
                print(f"[OK] {internal_id}")

            except Exception as e:
                row["status"] = "NG"
                row["message"] = f"{type(e).__name__}: {e}"
                print(f"[NG] {internal_id} -> {row['message']}")

            results.append(row)

    finally:
        pd.DataFrame(results).to_csv(log_path, index=False, encoding="utf-8-sig")
        print(f"\n日志已保存: {log_path}")
        # driver.quit()  # 需要自动关闭浏览器就取消注释


if __name__ == "__main__":
    main()
