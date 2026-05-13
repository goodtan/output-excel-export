import os
import sys
import time
import traceback

from openpyxl import load_workbook, Workbook
from playwright.sync_api import sync_playwright


if getattr(sys, "frozen", False):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


INPUT_EXCEL = os.path.join(app_dir(), "input.xlsx")
OUTPUT_EXCEL = os.path.join(app_dir(), "output.xlsx")
USER_DATA_DIR = os.path.join(app_dir(), "chrome-user-data")


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

    if contract_index is None:
        print("Excel 没找到合同编号列")
        return []

    tasks = []

    for row in ws.iter_rows(min_row=2):
        contract_no = row[contract_index].value
        if not contract_no:
            continue

        contract_no = str(contract_no).strip()

        detail_url = ""
        if url_index is not None:
            detail_url = row[url_index].value

        if not detail_url:
            print(f"{contract_no} 没有详情URL，跳过")
            continue

        tasks.append({
            "contract_no": contract_no,
            "detail_url": str(detail_url).strip(),
            "name": "",
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
        sign_btn = page.locator("button.ant-switch").filter(has_text="签入").first
        sign_btn.click(timeout=8000)
        time.sleep(2)
        print("已点击签入")
    except Exception:
        print("没有找到签入按钮，可能已经签入，跳过")


def get_name_from_page(page):
    try:
        body_text = page.locator("body").inner_text(timeout=5000)

        for line in body_text.splitlines():
            line = line.strip()
            if "承租人" in line and "性别" in line:
                return (
                    line.replace("承租人", "")
                    .replace("：", "")
                    .replace(":", "")
                    .strip()
                )
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
        phone_text = page.locator(".call-out").first.inner_text(timeout=5000)
        return phone_text.replace("\n", " ").strip()
    except Exception:
        return ""


def select_ant_option(page, input_id, option_text):
    select_root = page.locator(
        f"input#{input_id}"
    ).locator(
        "xpath=ancestor::div[contains(@class,'ant-select')]"
    )

    select_root.click(timeout=8000)
    time.sleep(0.5)

    page.locator(
        ".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option",
        has_text=option_text,
    ).last.click(timeout=8000)

    time.sleep(0.5)


def select_first_option(page, input_id):
    select_root = page.locator(
        f"input#{input_id}"
    ).locator(
        "xpath=ancestor::div[contains(@class,'ant-select')]"
    )

    select_root.click(timeout=8000)
    time.sleep(0.5)

    page.locator(
        ".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option"
    ).first.click(timeout=8000)

    time.sleep(0.5)


def click_call_btn(page):
    page.locator(".call-out img[src*='contractMakeCall']").first.click(timeout=10000)
    time.sleep(2)


def select_outbound_number(page):
    try:
        select_ant_option(page, "rc_select_3", "济南")
    except Exception:
        print("外显号码选择失败，跳过")


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


def submit_form(page):
    page.locator("button.ant-btn-primary:has-text('提 交')").first.click(timeout=10000)
    time.sleep(2)


def process_case(page, task):
    contract_no = task["contract_no"]
    detail_url = task["detail_url"]

    print(f"开始处理：{contract_no}")

    page.goto(detail_url, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=30000)
    time.sleep(3)

    sign_in(page)

    name = get_name_from_page(page)
    phone = get_real_phone(page)

    click_call_btn(page)
    select_outbound_number(page)

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
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            channel="chrome",
            headless=False,
            args=["--start-maximized"],
            viewport=None,
        )

        page = context.pages[0] if context.pages else context.new_page()

        print("=" * 50)
        print("第一次运行请手动登录系统")
        print("登录完成后按回车继续")
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
                    "name": task.get("name", ""),
                    "phone": "",
                    "status": "失败",
                    "error": str(e),
                }

            results.append(result)
            save_results(results)

        context.close()

    print("\n全部完成")
    print(f"结果文件：{OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
