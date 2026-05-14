import os
import sys
import time
import traceback

from openpyxl import load_workbook, Workbook
from playwright.sync_api import sync_playwright


INPUT_EXCEL = "input.xlsx"
OUTPUT_EXCEL = "output.xlsx"

USE_EXISTING_CHROME = True
CDP_URL = "http://127.0.0.1:9222"


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


INPUT_EXCEL = os.path.join(app_dir(), INPUT_EXCEL)
OUTPUT_EXCEL = os.path.join(app_dir(), OUTPUT_EXCEL)


def read_excel():
    wb = load_workbook(INPUT_EXCEL)
    ws = wb.active

    headers = {}
    for idx, cell in enumerate(ws[1]):
        if cell.value:
            headers[str(cell.value).strip()] = idx

    print("当前识别到的表头：", list(headers.keys()))

    contract_keys = ["合同编号", "合同号", "contractNo", "单号"]
    url_keys = ["详情URL", "详情页URL", "URL", "url", "链接", "详情页链接"]

    contract_index = None
    url_index = None

    for key in contract_keys:
        if key in headers:
            contract_index = headers[key]
            break

    for key in url_keys:
        if key in headers:
            url_index = headers[key]
            break

    if contract_index is None or url_index is None:
        print("Excel 必须包含：合同编号、详情URL")
        return []

    tasks = []

    for row in ws.iter_rows(min_row=2):
        contract_no = row[contract_index].value
        detail_url = row[url_index].value

        if not contract_no or not detail_url:
            continue

        tasks.append({
            "contract_no": str(contract_no).strip(),
            "detail_url": str(detail_url).strip(),
        })

    print(f"读取到 {len(tasks)} 条数据")
    return tasks


def save_results(results):
    wb = Workbook()
    ws = wb.active
    ws.title = "结果"
    ws.append(["合同编号", "姓名", "电话号码", "状态", "错误信息"])

    for item in results:
        ws.append([
            item.get("contract_no", ""),
            item.get("name", ""),
            item.get("phone", ""),
            item.get("status", ""),
            item.get("error", ""),
        ])

    wb.save(OUTPUT_EXCEL)


def sign_in(page):
    try:
        btn = page.locator("button.ant-switch").filter(has_text="签入").first
        btn.click(force=True, timeout=10000)
        print("已点击签入")
        time.sleep(3)
    except Exception as e:
        print("签入失败或已经签入，跳过：", e)


def change_current_status(page):
    try:
        page.wait_for_function(
            """
            () => {
                const el = document.querySelector('.current-status-value')
                return el && !el.className.includes('disabled')
            }
            """,
            timeout=15000
        )

        page.locator(".current-status-value").first.click(force=True, timeout=10000)
        time.sleep(1)

        page.locator(
            ".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option",
            has_text="空闲",
        ).last.click(force=True, timeout=10000)

        time.sleep(1)
        print("已切换状态为空闲")

    except Exception as e:
        print("切换状态失败，跳过：", e)


def select_outbound_number(page):
    try:
        page.locator(".dial-caller-select").first.wait_for(state="visible", timeout=15000)
        page.locator(".dial-caller-select").first.click(force=True, timeout=10000)
        time.sleep(1)

        page.locator(
            ".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option"
        ).first.click(force=True, timeout=10000)

        time.sleep(1)
        print("已选择外显号码")

    except Exception as e:
        print("选择外显号码失败：", e)
        raise


def get_name_from_page(page):
    try:
        body = page.locator("body").inner_text(timeout=5000)
        for line in body.splitlines():
            line = line.strip()
            if "承租人" in line and "性别" in line:
                return line
    except Exception:
        pass
    return ""


def get_real_phone(page):
    try:
        page.locator("span.show.toggle-des").first.click(timeout=5000)
        time.sleep(1)
    except Exception:
        pass

    try:
        text = page.locator(".call-out").first.inner_text(timeout=5000)
        return text.replace("\n", " ").strip()
    except Exception:
        return ""


def click_call_btn(page):
    page.locator(".call-out img[src*='contractMakeCall']").first.click(
        force=True,
        timeout=10000,
    )
    time.sleep(2)
    print("已点击拨打")


def hang_up(page):
    try:
        page.locator("button.call-button:has-text('挂断')").first.click(
            force=True,
            timeout=10000,
        )
    except Exception:
        page.locator("button.call-button").first.click(
            force=True,
            timeout=10000,
        )

    time.sleep(1)
    print("已挂断")


def select_ant_option(page, input_id, option_text):
    root = page.locator(f"input#{input_id}").locator(
        "xpath=ancestor::div[contains(@class,'ant-select')]"
    )

    root.click(force=True, timeout=8000)
    time.sleep(0.5)

    page.locator(
        ".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option",
        has_text=option_text,
    ).last.click(force=True, timeout=8000)

    time.sleep(0.5)


def select_first_option(page, input_id):
    root = page.locator(f"input#{input_id}").locator(
        "xpath=ancestor::div[contains(@class,'ant-select')]"
    )

    root.click(force=True, timeout=8000)
    time.sleep(0.5)

    page.locator(
        ".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option"
    ).first.click(force=True, timeout=8000)

    time.sleep(0.5)


def submit_form(page):
    page.locator("button.ant-btn-primary:has-text('提 交')").first.click(
        force=True,
        timeout=10000,
    )
    time.sleep(2)
    print("已提交")


def process_case(page, task):
    contract_no = task["contract_no"]
    detail_url = task["detail_url"]

    print(f"开始处理：{contract_no}")

    page.goto(detail_url, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=30000)
    time.sleep(3)

    sign_in(page)

    change_current_status(page)

    select_outbound_number(page)

    name = get_name_from_page(page)
    phone = get_real_phone(page)

    click_call_btn(page)

    time.sleep(3)

    hang_up(page)

    select_ant_option(page, "riskType", "失联")

    select_first_option(page, "contactResult")

    submit_form(page)

    print(f"完成：{contract_no}")

    return {
        "contract_no": contract_no,
        "name": name,
        "phone": phone,
        "status": "成功",
        "error": "",
    }


def get_page(playwright):
    if USE_EXISTING_CHROME:
        browser = playwright.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = context.pages[0] if context.pages else context.new_page()
        return browser, context, page

    context = playwright.chromium.launch_persistent_context(
        user_data_dir=os.path.join(app_dir(), "chrome-user-data"),
        channel="chrome",
        headless=False,
        args=["--start-maximized"],
        viewport=None,
    )

    page = context.pages[0] if context.pages else context.new_page()
    return None, context, page


def main():
    if not os.path.exists(INPUT_EXCEL):
        print(f"没找到 input.xlsx：{INPUT_EXCEL}")
        return

    tasks = read_excel()

    if not tasks:
        print("Excel 没有数据")
        return

    results = []

    with sync_playwright() as p:
        browser, context, page = get_page(p)

        print("=" * 50)
        print("请确认 Chrome 已登录系统")
        print("确认后按回车继续")
        print("=" * 50)
        input()

        for index, task in enumerate(tasks, start=1):
            print(f"\n[{index}/{len(tasks)}]")

            try:
                result = process_case(page, task)
            except Exception as e:
                print("处理失败")
                print(traceback.format_exc())

                result = {
                    "contract_no": task.get("contract_no", ""),
                    "name": "",
                    "phone": "",
                    "status": "失败",
                    "error": str(e),
                }

            results.append(result)
            save_results(results)

        if browser:
            browser.close()
        else:
            context.close()

    print("\n全部完成")
    print(f"结果文件：{OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
