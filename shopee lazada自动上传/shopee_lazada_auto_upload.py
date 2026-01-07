import os
import re
import csv
import time
from pathlib import Path
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    NoAlertPresentException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
)


# ================== 配置区域 ==================

NETSUITE_LOGIN_URL = (
    "https://6806569.app.netsuite.com/app/login/secure/enterpriselogin.nl?"
    "c=6806569&redirect=%2Fapp%2Faccounting%2Ftransactions%2Fcustpymt.nl"
    "%3Fid%3D6402549%26whence%3D&whence="
)

IMPORT_URL = (
    "https://6806569.app.netsuite.com/app/setup/assistants/nsimport/"
    "importassistant.nl?recid=148&new=T&whence=&siaT=1765412585636&"
    "siaWhc=%2Fapp%2Faccounting%2Ftransactions%2Fcustpymt.nl&siaNv=ct2"
)

QUEUE_URL = (
    "https://6806569.app.netsuite.com/app/site/hosting/scriptlet.nl?"
    "script=750&deploy=1&whence=&siaT=1765415091644&"
    "siaWhc=%2Fapp%2Fsetup%2Fassistants%2Fnsimport%2Fimportassistant.nl&siaNv=ct2"
)

CUSTPYMT_URL = (
    "https://6806569.app.netsuite.com/app/accounting/transactions/"
    "transactionlist.nl?Transaction_TYPE=CustPymt&whence=&siaT=1765415706543&"
    "siaWhc=%2Fapp%2Fsite%2Fhosting%2Fscriptlet.nl&siaNv=ct3"
)

UPLOAD_DIR = r"C:\Users\mitsu\OneDrive\デスクトップ\shopee lazada自动上传\自动上传"
LOG_FILE = r"C:\Users\mitsu\OneDrive\デスクトップ\shopee lazada自动上传\upload_log.csv"

FILE_EXTENSIONS = {".csv"}


# ================== 基础工具函数 ==================

def init_driver():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver


def scroll_into_view(driver, element, center=True):
    if center:
        driver.execute_script(
            "arguments[0].scrollIntoView({behavior:'smooth', block:'center'});",
            element,
        )
    else:
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
    time.sleep(0.5)


def scroll_to_top(driver):
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.3)


def click_blank_area(driver):
    """
    目前主要用于兜底场景（基本不再依赖它触发 onchange）
    """
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.click()
        return
    except Exception:
        pass

    for css in ["#div__body", "#main_form", "html"]:
        try:
            elem = driver.find_element(By.CSS_SELECTOR, css)
            elem.click()
            return
        except Exception:
            continue


def wait_for_step1_page(driver):
    wait = WebDriverWait(driver, 20)
    elem = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "//*[contains(text(),'CSVファイルのスキャンとアップロード')]")
        )
    )
    scroll_into_view(driver, elem, center=True)
    return wait


def set_char_encoding_utf8(driver, wait):
    char_input = wait.until(
        EC.element_to_be_clickable((By.NAME, "inpt_charencoding"))
    )
    scroll_into_view(driver, char_input, center=True)
    char_input.click()
    time.sleep(0.3)

    for _ in range(3):
        char_input.send_keys(Keys.ARROW_UP)
        time.sleep(0.1)

    char_input.send_keys(Keys.ENTER)
    time.sleep(0.5)


def list_all_files():
    p = Path(UPLOAD_DIR)
    if not p.exists():
        return []
    return sorted(
        [f for f in p.iterdir() if f.is_file() and f.suffix.lower() in FILE_EXTENSIONS],
        key=lambda x: x.name,
    )


def ensure_log_dir():
    log_dir = os.path.dirname(LOG_FILE)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)


def load_uploaded_filenames():
    ensure_log_dir()
    if not os.path.exists(LOG_FILE):
        return set()

    uploaded = set()
    with open(LOG_FILE, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            filename = row.get("filename")
            if filename:
                uploaded.add(filename)
    return uploaded


def get_next_file():
    all_files = list_all_files()
    uploaded = load_uploaded_filenames()
    for f in all_files:
        if f.name not in uploaded:
            return f
    return None


def parse_filename(filename):
    basename = Path(filename).name
    pattern = re.compile(
        r"^(?P<platform>shopee|lazada)-"
        r"(?P<country>[A-Z]{2})-"
        r"(?P<shop>C\d{6})-"
        r"(?P<year>\d{4})-"
        r"(?P<md>\d{4})",
        re.IGNORECASE,
    )
    m = pattern.search(basename)
    if not m:
        return None, None, None, None

    platform = m.group("platform").lower()
    country = m.group("country").upper()
    shop = m.group("shop").upper()
    year = int(m.group("year"))
    md = m.group("md")

    try:
        month = int(md[:2])
        day = int(md[2:])
        dt = datetime(year, month, day)
        date_norm = dt.strftime("%Y-%m-%d")
    except Exception:
        date_norm = f"{year}-{md}"

    return platform, country, shop, date_norm


def append_log(filename, country, shop, date_str):
    ensure_log_dir()
    file_exists = os.path.exists(LOG_FILE)

    with open(LOG_FILE, "a", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["timestamp", "filename", "country", "shop", "date"])
        if not file_exists:
            writer.writeheader()

        writer.writerow({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "filename": filename,
            "country": country or "",
            "shop": shop or "",
            "date": date_str or "",
        })


def upload_file_step1(driver, wait, filepath: Path):
    full_path = str(filepath.resolve())
    file_input = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
    )
    scroll_into_view(driver, file_input, center=True)
    file_input.send_keys(full_path)
    time.sleep(1.0)


# -----------------
# 🔥 click_next：共通的「次へ >」按钮
# -----------------

def click_next(driver, wait):
    print("准备点击『次へ >』按钮 ...")

    try:
        next_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "next"))
        )
    except TimeoutException:
        next_btn = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//input[@type='button' and contains(@value,'次へ')]"
                    " | //button[contains(normalize-space(),'次へ')]",
                )
            )
        )

    scroll_into_view(driver, next_btn, center=True)

    try:
        next_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", next_btn)

    time.sleep(3)


# ================== Step2：インポート・オプション ==================

def handle_import_options_step2(driver, wait):
    """展开アドバンスト・オプション，选择 Custom Form 2，然后 Next"""
    try:
        title = wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "//*[contains(text(),'インポート・オプション')]")
            )
        )
        scroll_into_view(driver, title, center=True)
    except Exception:
        pass

    # 展开「アドバンスト・オプション」
    try:
        adv_row = driver.find_element(By.ID, "tr_fldarr_1")
        if not adv_row.is_displayed():
            adv_label = wait.until(
                EC.element_to_be_clickable((By.ID, "label_fldarr_1"))
            )
            scroll_into_view(driver, adv_label, center=True)
            adv_label.click()
            time.sleep(0.5)
    except Exception as e:
        print("警告：展开『アドバンスト・オプション』时出错（可能已展开）：", e)

    # 选择 Custom Form
    customform_select = wait.until(
        EC.element_to_be_clickable((By.ID, "customform"))
    )
    scroll_into_view(driver, customform_select, center=True)
    sel = Select(customform_select)
    try:
        sel.select_by_visible_text("Standard 订单后续中间表 Form 2")
    except Exception:
        print("按文本选择失败，尝试选择第2个选项(index=1)")
        sel.select_by_index(1)
    time.sleep(0.5)

    # 这里用 10 秒等待的 click_next
    click_next(driver, WebDriverWait(driver, 10))


# ================== Step4：字段映射辅助函数 ==================

def _expand_label_amount_variants(label: str) -> list[str]:
    """
    扩展“金额”相关字段的字形：额/額 都能匹配。
    例如：退款金额 -> [退款金额, 退款金額]
    """
    variants = {label}

    # 优先按“金额/金額”整词替换
    if "金额" in label:
        variants.add(label.replace("金额", "金額"))
    if "金額" in label:
        variants.add(label.replace("金額", "金额"))

    # 兜底：单字替换（防止出现非整词场景）
    if "额" in label:
        variants.add(label.replace("额", "額"))
    if "額" in label:
        variants.add(label.replace("額", "额"))

    return sorted(variants)


def click_tree_node_by_label(
    driver,
    tree_div_id: str,
    label: str,
    timeout: int = 20,
    retries: int = 3,
):
    """
    在字段树（左/右）里通过文字点击节点（增强版）：
    - 对“金额”字段自动兼容 额/額
    - 支持 alttext/text 的精确与包含匹配（NetSuite 有时会带前后缀）
    - 自动重试 + JS click 兜底 + 处理 stale
    tree_div_id: 'filecoltree_b'（左）或 'ltfieldtree_b'（右）
    """
    driver.switch_to.default_content()
    wait = WebDriverWait(driver, timeout)

    base = f"//div[@id='{tree_div_id}']"
    labels = _expand_label_amount_variants(label)

    conds = []
    for lb in labels:
        lb_esc = lb.replace("'", "\'")
        conds.append(
            f"(@alttext and (normalize-space(@alttext)='{lb_esc}' or contains(normalize-space(@alttext),'{lb_esc}')))"
        )
        conds.append(
            f"(normalize-space(text())='{lb_esc}' or contains(normalize-space(text()),'{lb_esc}'))"
        )

    node_xpath = f"{base}//*[{' or '.join(conds)}]"

    last_err = None
    for _ in range(retries):
        try:
            elem = wait.until(EC.presence_of_element_located((By.XPATH, node_xpath)))
            scroll_into_view(driver, elem, center=True)

            # 尽量点击更“可点”的祖先元素，避免点到纯文本节点
            try:
                clickable = elem.find_element(
                    By.XPATH, "./ancestor-or-self::*[self::a or self::span or self::div][1]"
                )
            except Exception:
                clickable = elem

            wait.until(EC.element_to_be_clickable(clickable))

            try:
                clickable.click()
            except (ElementClickInterceptedException, ElementNotInteractableException):
                driver.execute_script("arguments[0].click();", clickable)

            time.sleep(0.3)
            driver.switch_to.default_content()
            return

        except StaleElementReferenceException as e:
            last_err = e
            time.sleep(0.4)
            driver.switch_to.default_content()
            continue
        except Exception as e:
            last_err = e
            time.sleep(0.4)
            driver.switch_to.default_content()
            continue

    raise RuntimeError(
        f"点击树节点失败 tree={tree_div_id}, label={label}, candidates={labels}, err={last_err}"
    )


def click_middle_row_by_label(driver, label: str, timeout: int = 10):
    """
    在中间选择框里，通过左侧字段名点击对应行。
    同样对“金额”字段自动兼容 额/額。
    结构示例：
      <span title="xxx: 付款金额">付款金额</span>
    """
    wait = WebDriverWait(driver, timeout)
    labels = _expand_label_amount_variants(label)

    conds = []
    for lb in labels:
        lb_esc = lb.replace("'", "\'")
        conds.append(f"normalize-space(text())='{lb_esc}'")
        conds.append(f"contains(normalize-space(text()),'{lb_esc}')")

    xpath = (
        "//div[@id='mapperpane']//tr"
        f"[.//span[{ ' or '.join(conds) }]]"
    )

    row = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    scroll_into_view(driver, row, center=True)
    try:
        row.click()
    except (ElementClickInterceptedException, ElementNotInteractableException):
        driver.execute_script("arguments[0].click();", row)
    time.sleep(0.3)



# ================== Step4：フィールド・マッピング ==================

def handle_field_mapping_step4(driver, wait):
    """
    Step4: フィールド・マッピング
    依次为「付款金额」「退款金额」「账单金额」建立映射：
      左（你的字段）→ 中（对应行）→ 右（NetSuiteフィールド）
    然后点击『次へ >』进入保存&実行页面
    """
    print("进入 Step4：フィールド・マッピング（金额字段映射）")

    driver.switch_to.default_content()
    wait.until(EC.presence_of_element_located((By.ID, "mapperpane")))
    scroll_to_top(driver)

    try:
        title = driver.find_element(
            By.XPATH, "//*[contains(text(),'フィールド・マッピング')]"
        )
        scroll_into_view(driver, title, center=True)
    except Exception:
        pass

    fields = ["付款金额", "退款金额", "账单金额"]

    for label in fields:
        print(f"  正在映射字段：{label}")

        # 1) 左侧 あなたのフィールド（filecoltree_b）
        try:
            click_tree_node_by_label(driver, "filecoltree_b", label)
        except Exception as e:
            print(f"    [警告] 左侧字段『{label}』点击失败：{e}")
            continue

        # 2) 中间选择框对应行
        try:
            driver.switch_to.default_content()
            click_middle_row_by_label(driver, label)
        except Exception as e:
            print(f"    [警告] 中间选择框中『{label}』行点击失败（可能已自动选中）：{e}")
            driver.switch_to.default_content()

        # 3) 右侧 NetSuiteフィールド（ltfieldtree_b）
        try:
            click_tree_node_by_label(driver, "ltfieldtree_b", label)
        except Exception as e:
            print(f"    [错误] 右侧字段『{label}』点击失败：{e}")
            driver.switch_to.default_content()
            continue

        time.sleep(0.5)

    driver.switch_to.default_content()
    time.sleep(0.3)  # 映射结束后稍微等一下就点次へ

    print("  字段映射结束，点击『次へ >』进入 Step5...")
    # 这里也改为 10 秒等待
    click_next(driver, WebDriverWait(driver, 10))
    print("已从 Step4 跳转，开始加载 Step5。")


# ================== Step5：保存 & 实行（点「実行」+ 处理弹窗） ==================

def handle_save_and_run_step5(driver, wait):
    """
    Step5: マッピングを保存してインポートを開始

    直接触发隐藏按钮 finish（= 実行），然后处理 JS 确认弹窗。
    """
    print("开始处理 Step5：保存 & 实行（直接触发『実行』）...")
    driver.switch_to.default_content()

    # 1) 确认已经到了 Step5 页面（标题）
    try:
        long_wait = WebDriverWait(driver, 60)
        title = long_wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "//*[contains(text(),'マッピングを保存してインポートを開始')"
                    " or contains(text(),'インポートを開始')]",
                )
            )
        )
        scroll_into_view(driver, title, center=True)
        print("  已检测到 Step5 标题区域。")
    except TimeoutException:
        print("  [警告] 60 秒内未检测到 Step5 标题，可能停在错误页面或其他页面。")
        print("  当前 URL:", driver.current_url)

    # 2) 等待隐藏的 finish 按钮（对应菜单里的「実行」）
    try:
        finish_btn = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.ID, "finish"))
        )
        print("  已找到隐藏按钮 id='finish'（メニューの「実行」）。")
    except TimeoutException:
        print("  [错误] 40 秒内没有找到 id='finish' 的按钮，无法执行导入。")
        return
    except Exception as e:
        print("  [错误] 查找 id='finish' 按钮时发生异常：", e)
        return

    # 3) 通过 JS 触发『実行』按钮
    try:
        try:
            scroll_into_view(driver, finish_btn, center=False)
        except Exception:
            pass

        driver.execute_script("arguments[0].click();", finish_btn)
        print("  已通过 JS 触发『実行』(finish) 按钮，等待确认弹窗 ...")
    except Exception as e:
        print("  [错误] 点击『実行』(finish) 按钮过程出错：", e)
        return

    # 4) 处理浏览器确认弹窗（点「确定 / OK」）
    try:
        alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
        print("  检测到确认弹窗，点击『确定』...")
        alert.accept()
        time.sleep(2)
    except TimeoutException:
        print("  10 秒内没有出现确认弹窗，可能当前设置不再提示，直接继续。")
    except NoAlertPresentException:
        print("  未检测到确认弹窗（NoAlertPresent），继续流程。")
    except Exception as e:
        print("  处理确认弹窗时发生异常：", e)

    print("  Step5『実行』及确认已完成，等待 NetSuite 处理导入任务 ...")
    time.sleep(3)


# ================== 队列页面：轮询 Submit ==================

def wait_and_submit_queue(driver, max_retries=10):
    """轮询 scriptlet 页面，看到『当前还有 xxx 条记录待处理』且有 Submit 时点击"""
    print("等待 60 秒后，跳转到队列监控页面 ...")
    time.sleep(60)

    for attempt in range(1, max_retries + 1):
        print(f"[队列监控] 第 {attempt} 次检查 ...")
        driver.get(QUEUE_URL)
        time.sleep(2)

        page_text = driver.page_source
        m = re.search(r"当前还有\s*(\d+)\s*条记录待处理", page_text)
        pending = None
        if m:
            pending = int(m.group(1))
            print(f"  检测到提示：当前还有 {pending} 条记录待处理")
        else:
            print("  未找到提示文本")

        try:
            submit_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//input[(@type='button' or @type='submit') and "
                        "(translate(@value,'SUBMIT','submit')='submit')]"
                        " | //button[translate(normalize-space(),'SUBMIT','submit')='submit']",
                    )
                )
            )
            scroll_into_view(driver, submit_btn, center=True)
            has_button = True
        except TimeoutException:
            print("  未找到 Submit 按钮")
            has_button = False
            submit_btn = None

        if pending is not None and pending > 0 and has_button:
            print("  条件满足，点击 Submit。")
            try:
                submit_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", submit_btn)
            time.sleep(3)
            return

        print("  条件不满足，60 秒后刷新重试 ...")
        time.sleep(60)

    print("警告：队列监控超过最大重试次数。")


# ================== 🔥 新增：直接操纵 NetSuite 下拉（searchid / Transaction_NAME） ==================

def set_netsuite_dropdown_by_text(driver, data_name, label, partial=False):
    """
    利用 ns-dropdown 的 data-options，根据显示文本（label）设置下拉值，
    并触发 NetSuite 的 onchange（从而刷新页面）。

    data_name: ns-dropdown 的 data-name，比如 "searchid", "Transaction_NAME"
    label    : 要选中的显示文字（视图名或店铺整行文字的一部分）
    partial  : True=包含匹配；False=全等匹配
    """
    script = r"""
    var dataName = arguments[0];
    var label = arguments[1];
    var partial = arguments[2];

    var dropdowns = document.querySelectorAll('div.ns-dropdown');
    var ddDiv = null;
    for (var i = 0; i < dropdowns.length; i++) {
        if (dropdowns[i].getAttribute('data-name') === dataName) {
            ddDiv = dropdowns[i];
            break;
        }
    }
    if (!ddDiv) {
        return { ok: false, error: 'dropdown_not_found', dataName: dataName };
    }

    var optionsJson = ddDiv.getAttribute('data-options');
    var opts;
    try {
        opts = JSON.parse(optionsJson);
    } catch (e) {
        return { ok: false, error: 'json_parse_failed', detail: '' + e, raw: optionsJson };
    }

    var match = null;
    for (var j = 0; j < opts.length; j++) {
        var t = opts[j].text;
        if (!t) continue;
        if (partial) {
            if (t.indexOf(label) !== -1) {
                match = opts[j];
                break;
            }
        } else {
            if (t === label) {
                match = opts[j];
                break;
            }
        }
    }

    if (!match) {
        return { ok: false, error: 'option_not_found', label: label };
    }

    var hiddenList = document.getElementsByName(dataName);
    if (!hiddenList || !hiddenList.length) {
        return { ok: false, error: 'hidden_input_not_found', dataName: dataName, value: match.value };
    }
    var hidden = hiddenList[0];
    hidden.value = match.value;

    var inputName = 'inpt_' + dataName;
    var inputList = document.getElementsByName(inputName);
    if (inputList && inputList.length) {
        var inp = inputList[0];
        inp.value = match.text;
        if (window.getDropdown) {
            try {
                var dd = getDropdown(inp);
                if (dd && dd.setValue) {
                    dd.setValue(match.value);
                }
            } catch(e) {
                // ignore
            }
        }
    }

    if (typeof hidden.onchange === 'function') {
        hidden.onchange();
    }

    return { ok: true, value: match.value, text: match.text, dataName: dataName };
    """

    result = driver.execute_script(script, data_name, label, partial)
    if not result or not result.get("ok"):
        raise RuntimeError(f"设置下拉框 {data_name} 失败: {result}")
    return result


def ensure_filter_expanded(driver, wait):
    """
    用 aria-controls / aria-expanded 判断『フィルター』区域是否展开，
    未展开则点击一次。
    """
    print("检查『フィルター』区域是否已展开...")

    try:
        header = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "[aria-controls='uir-filters-body']")
            )
        )
    except TimeoutException:
        print("警告：没有找到控制『フィルター』区域的 header 元素，暂时跳过展开判断。")
        return

    try:
        expanded = header.get_attribute("aria-expanded")
    except StaleElementReferenceException:
        header = driver.find_element(By.CSS_SELECTOR, "[aria-controls='uir-filters-body']")
        expanded = header.get_attribute("aria-expanded")

    if expanded and expanded.lower() == "true":
        print("『フィルター』已经是展开状态。")
        return

    print("『フィルター』目前是收起状态，点击一次将其展开。")
    scroll_into_view(driver, header, center=True)
    header.click()

    # 等 aria-expanded 变为 true
    try:
        WebDriverWait(driver, 10).until(
            lambda d: d.find_element(
                By.CSS_SELECTOR, "[aria-controls='uir-filters-body']"
            ).get_attribute("aria-expanded") == "true"
        )
        print("『フィルター』区域已成功展开。")
    except TimeoutException:
        print("警告：点击『フィルター』后 aria-expanded 没有变为 true，不过继续后续操作。")

    time.sleep(1.5)


def apply_view_and_filter_by_shop(driver, wait, shop_code):
    """
    入金列表设置视图 + 根据店铺代码过滤『名前』：
      ① searchid（表示）= FB_トランザクション【BToB】経理
      ② 展开フィルター
      ③ Transaction_NAME（名前）用 shop_code 部分匹配
    """

    # ----------- Step1: 设置「表示」视图 ----------

    view_text = "FB_トランザクション【BToB】経理"
    try:
        print(f"设置『表示』为：{view_text}")
        set_netsuite_dropdown_by_text(driver, "searchid", view_text, partial=False)
        time.sleep(4)
    except Exception as e:
        print("警告：设置『表示』视图失败：", e)

    # ----------- Step2: 确保『フィルター』展开 ----------

    ensure_filter_expanded(driver, wait)

    # ----------- Step3: 设置『名前』过滤 ----------

    try:
        print(f"设置『名前』包含店铺代码：{shop_code}")
        # Transaction_NAME 的 text 通常为：
        #   C000126 アマゾンジャパン【BToB専用】
        # 这里用 partial=True，只要 text 中包含 C000126 即可
        set_netsuite_dropdown_by_text(driver, "Transaction_NAME", shop_code, partial=True)
        time.sleep(4)
    except Exception as e:
        print("警告：设置『名前』过滤失败：", e)


# ================== 入金列表：检查记录 ==================

def check_transaction_row_exists(driver, target_date_str):
    """检查当前列表是否存在 日付=target_date_str 且 メモ 为空 的行"""
    print(f"检查日期 {target_date_str} 的记录是否出现...")

    date_cells = driver.find_elements(
        By.XPATH, f"//*[normalize-space()='{target_date_str}']"
    )

    for cell in date_cells:
        try:
            row = cell.find_element(By.XPATH, "./ancestor::tr[1]")
        except NoSuchElementException:
            continue

        try:
            memo_cell = row.find_element(
                By.XPATH,
                ".//td[contains(@data-label,'メモ') or contains(@aria-label,'メモ')]",
            )
            memo_text = memo_cell.text.strip()
            if memo_text == "":
                print("找到目标记录（メモ为空）。")
                return True
        except NoSuchElementException:
            print("找到目标记录（没有メモ列）。")
            return True

    return False


def wait_for_transaction_in_list(driver, shop_code, date_str, max_retries=60):
    """轮询入金列表，直到出现指定店铺 & 日期记录"""

    if not date_str:
        print("没有解析到日期，跳过入金列表检查。")
        return

    target_date = date_str.replace("-", "/")

    for attempt in range(1, max_retries + 1):
        print(f"[入金列表] 第 {attempt} 次检查 ...")
        driver.get(CUSTPYMT_URL)
        wait = WebDriverWait(driver, 20)

        # 应用过滤（视图 + 店铺）
        apply_view_and_filter_by_shop(driver, wait, shop_code)

        # 等待刷新后再找记录
        time.sleep(2)

        if check_transaction_row_exists(driver, target_date):
            print(f"  🎉 找到日期为 {target_date} 且メモ为空的记录！")
            return

        print("  未找到目标记录，10 秒后刷新重试 ...")
        time.sleep(10)

    print("⚠ 警告：入金列表监控超过最大重试次数。")


# ================== 主流程 ==================

def main():
    driver = init_driver()

    try:
        # 登录
        driver.get(NETSUITE_LOGIN_URL)
        print("已打开 NetSuite 登录页面。")
        print("请在浏览器中手动登录（包含 2FA），完成后回到此窗口。")
        input("登录完成后按 Enter 继续...")

        # 循环所有未处理文件
        while True:
            next_file = get_next_file()
            if not next_file:
                print("没有更多未处理的文件，程序结束。")
                break

            print("\n============================")
            print(f"开始处理文件: {next_file.name}")

            platform, country, shop, date_str = parse_filename(next_file.name)
            print("解析结果 ->", platform, country, shop, date_str)

            # Step1
            driver.get(IMPORT_URL)
            wait = wait_for_step1_page(driver)

            set_char_encoding_utf8(driver, wait)
            upload_file_step1(driver, wait, next_file)
            append_log(next_file.name, country, shop, date_str)
            click_next(driver, WebDriverWait(driver, 10))

            # Step2
            wait = WebDriverWait(driver, 20)
            handle_import_options_step2(driver, wait)

            # Step4 字段映射
            wait = WebDriverWait(driver, 20)
            handle_field_mapping_step4(driver, wait)

            # Step5 保存 & 实行
            wait = WebDriverWait(driver, 20)
            handle_save_and_run_step5(driver, wait)

            # 队列 submit
            wait_and_submit_queue(driver, max_retries=60)

            # 入金列表确认
            if shop and date_str:
                wait_for_transaction_in_list(driver, shop, date_str, max_retries=60)
            else:
                print("未解析出 shop/date，跳过入金确认。")

            print(f"文件 {next_file.name} 处理完毕，进入下一份文件。")

    finally:
        print("流程结束。（调试阶段可先保留浏览器，确认行为后再打开 quit）")
        # driver.quit()


if __name__ == "__main__":
    main()
