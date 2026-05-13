import os
import sys
import time
import traceback
from urllib.parse import urlparse, parse_qs, unquote

from openpyxl import load_workbook, Workbook
from playwright.sync_api import sync_playwright


# =========================
# PyInstaller + Playwright
# =========================
if getattr(sys, "frozen", False):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"


# =========================
# 文件配置
# =========================
INPUT_EXCEL = "input.xlsx"
OUTPUT_EXCEL = "output.xlsx"

# Chrome 用户缓存目录
USER_DATA_DIR = "./chrome-user-data"


# =========================
# 读取 Excel
# =========================
def read_excel():
    wb = load_workbook(INPUT_EXCEL)
    ws = wb.active

    headers = {}

    for idx, cell in enumerate(ws[1]):
        headers[str(cell.value).strip()] = idx

    tasks = []

    for row in ws.iter_rows(min_row=2):
        contract_no = ""
        detail_url = ""

        if "合同编号" in headers:
            contract_no = row[headers["合同编号"]].value

        if "详情页URL" in headers:
            detail_url = row[headers["详情页URL"]].value

        if not detail_url:
            continue

        detail_url = str(detail_url).strip()

        if not contract_no:
            contract_no = parse_contract_no(detail_url)

        tasks.append({
            "contract_no": str(contract_no).strip(),
            "detail_url": detail_url,
            "name": parse_name(detail_url),
        })

    return tasks


# =========================
# 从 URL 获取合同号
# =========================
def parse_contract_no(url):
    try:
        query = parse_qs(urlparse(url).query)
        return query.get("contractNo", [""])[0]
    except Exception:
        return ""


# =========================
# 从 URL 获取姓名
# =========================
def parse_name(url):
    try:
        query = parse_qs(urlparse(url).query)
        return unquote(query.get("loanName", [""])[0])
    except Exception:
        return ""


# =========================
# 保存结果
# =========================
def save_results(results):
    wb = Workbook()
    ws = wb.active

    ws.title = "结果"

    ws.append([
        "合同编号",
        "姓名",
        "电话号码",
        "状态",
        "错误信息"
    ])

    for item in results:
        ws.append([
            item.get("contract_no", ""),
            item.get("name", ""),
            item.get("phone", ""),
            item.get("status", ""),
            item.get("error", ""),
        ])

    wb.save(OUTPUT_EXCEL)


# =========================
# 获取真实手机号
# =========================
def get_real_phone(page):
    try:
        # 点击眼睛
        eye_btn = page.locator("span.show.toggle-des").first
        eye_btn.click(timeout=5000)

        time.sleep(1)

    except Exception:
        pass

    try:
        phone_text = page.locator(".call-out").first.inner_text(timeout=5000)

        return phone_text.replace("\n", " ").strip()

    except Exception:
        return ""


# =========================
# 选择 ant 下拉
# =========================
def select_ant_option(page, input_id, option_text):
    select_root = page.locator(
        f"input#{input_id}"
    ).locator(
        "xpath=ancestor::div[contains(@class,'ant-select')]"
    )

    select_root.click(timeout=5000)

    time.sleep(1)

    option = page.locator(
        ".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option",
        has_text=option_text
    ).last

    option.click(timeout=5000)

    time.sleep(1)


# =========================
# 随便选一个 option
# =========================
def select_first_option(page, input_id):
    select_root = page.locator(
        f"input#{input_id}"
    ).locator(
        "xpath=ancestor::div[contains(@class,'ant-select')]"
    )

    select_root.click(timeout=5000)

    time.sleep(1)

    option = page.locator(
        ".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option"
    ).first

    option.click(timeout=5000)

    time.sleep(1)


# =========================
# 拨打电话
# =========================
def click_call_btn(page):
    call_btn = page.locator(
        ".call-out img[src*='contractMakeCall']"
    ).first

    call_btn.click(timeout=10000)

    time.sleep(2)


# =========================
# 选择外显号码
# =========================
def select_outbound_number(page):
    try:
        select_ant_option(page, "rc_select_3", "济南")
    except Exception:
        print("外显号码选择失败，跳过")


# =========================
# 挂断
# =========================
def hang_up(page):
    try:
        hangup_btn = page.locator(
            "button.call-button:has-text('挂断')"
        ).first

        hangup_btn.click(force=True, timeout=10000)

    except Exception:
        page.locator("button.call-button").first.click(
            force=True,
            timeout=10000
        )

    time.sleep(1)


# =========================
# 风险分类
# =========================
def set_risk_type(page):
    select_ant_option(page, "riskType", "失联")


# =========================
# 联络结果
# =========================
def set_contact_result(page):
    select_first_option(page, "contactResult")


# =========================
# 提交
# =========================
def submit_form(page):
    submit_btn = page.locator(
        "button.ant-btn-primary:has-text('提 交')"
    ).first

    submit_btn.click(timeout=10000)

    time.sleep(2)


# =========================
# 单条处理
# =========================
def process_case(page, task):
    contract_no = task["contract_no"]
    detail_url = task["detail_url"]
    name = task["name"]

    print(f"开始处理：{contract_no}")

    page.goto(detail_url)

    page.wait_for_load_state("networkidle")

    time.sleep(3)

    # 获取手机号
    phone = get_real_phone(page)

    # 点击拨打
    click_call_btn(page)

    # 选择外显号码
    select_outbound_number(page)

    # 等待 3 秒
    time.sleep(3)

    # 挂断
    hang_up(page)

    # 风险分类
    set_risk_type(page)

    # 联络结果
    set_contact_result(page)

    # 提交
    submit_form(page)

    print(f"完成：{contract_no}")

    return {
        "contract_no": contract_no,
        "name": name,
        "phone": phone,
        "status": "成功",
        "error": "",
    }


# =========================
# 主流程
# =========================
def main():
    tasks = read_excel()

    if not tasks:
        print("Excel 没有数据")
        return

    results = []

    with sync_playwright() as p:

        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,

            # 使用系统 Chrome
            channel="chrome",

            headless=False,

            args=[
                "--start-maximized"
            ],

            viewport=None
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

            # 每处理一条就保存一次
            save_results(results)

        context.close()

    print("\n全部完成")
    print(f"结果文件：{OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
