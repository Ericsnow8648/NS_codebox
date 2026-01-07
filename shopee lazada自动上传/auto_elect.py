import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

NETSUITE_LOGIN_URL = (
    "https://6806569.app.netsuite.com/app/login/secure/enterpriselogin.nl?"
    "c=6806569&redirect=%2Fapp%2Faccounting%2Ftransactions%2Fcustpymt.nl"
    "%3Fid%3D6402549%26whence%3D&whence="
)

CUSTPYMT_URL = (
    "https://6806569.app.netsuite.com/app/accounting/transactions/"
    "transactionlist.nl?Transaction_TYPE=CustPymt&whence=&siaT=1765415706543&"
    "siaWhc=%2Fapp%2Fsite%2Fhosting%2Fscriptlet.nl&siaNv=ct3"
)


def init_driver():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver


def set_netsuite_dropdown_by_text(driver, data_name, label, partial=False):
    """
    利用 ns-dropdown 的 data-options，根据显示文本（label）设置下拉值，
    并触发 NetSuite 的 onchange（从而刷新页面）。

    data_name: ns-dropdown 的 data-name，比如 "searchid", "Transaction_NAME"
    label    : 要选中的显示文字
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


def main():
    driver = init_driver()
    wait = WebDriverWait(driver, 20)

    try:
        # 1. 登录
        driver.get(NETSUITE_LOGIN_URL)
        print("已打开 NetSuite 登录页面。")
        print("请在浏览器中手动登录（包含 2FA），完成后回到此窗口。")
        input("登录完成后按 Enter 继续...")

        # 2. 打开入金列表页面
        driver.get(CUSTPYMT_URL)
        print("已打开入金列表页面。")

        # 3. 输入店铺代码
        while True:
            shop_code = input("\n请输入店铺编号（例如 C000126，按 Enter 使用默认）：").strip()
            if shop_code == "":
                shop_code = "C000126"   # 默认店铺
                print("使用默认店铺：", shop_code)
                break
            elif shop_code.upper().startswith("C") and len(shop_code) >= 5:
                shop_code = shop_code.upper()
                print("将使用店铺编号：", shop_code)
                break
            else:
                print("格式不正确，请重新输入（例如 C000126）。")

        # 视图名称保持不变
        view_text = "FB_トランザクション【BToB】経理"

        print("\n====== 开始设置视图 ======")
        set_netsuite_dropdown_by_text(driver, "searchid", view_text, partial=False)
        time.sleep(4)
        print("『表示』设置完成。\n")

        # 显示当前视图文字
        try:
            cur_view = wait.until(
                EC.presence_of_element_located((By.NAME, "inpt_searchid"))
            ).get_attribute("value")
            print("当前视图显示文字：", cur_view)
        except:
            pass

        print("\n====== 开始设置『名前』过滤 ======")
        set_netsuite_dropdown_by_text(driver, "Transaction_NAME", shop_code, partial=True)
        time.sleep(4)
        print("『名前』设置完成。\n")

        # 显示当前店铺文字
        try:
            cur_name = wait.until(
                EC.presence_of_element_located((By.NAME, "inpt_Transaction_NAME"))
            ).get_attribute("value")
            print("当前『名前』显示文字：", cur_name)
        except:
            pass

        print("\n=== 已完成过滤，请在页面查看是否正确显示该店铺交易 ===\n")

    finally:
        print("调试结束（浏览器保持打开）。")
        # driver.quit()


if __name__ == "__main__":
    main()
