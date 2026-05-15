import os
import sys
import time
import traceback
from datetime import datetime

from openpyxl import load_workbook, Workbook
from playwright.sync_api import sync_playwright


INPUT_EXCEL_NAME = "input.xlsx"
OUTPUT_EXCEL_NAME = "output.xlsx"
CDP_URL = "http://127.0.0.1:9222"


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


INPUT_EXCEL = os.path.join(app_dir(), INPUT_EXCEL_NAME)
OUTPUT_EXCEL = os.path.join(app_dir(), OUTPUT_EXCEL_NAME)


def pause_exit():
    print("\n按回车键退出窗口...")
    input()


def read_excel():
    wb = load_workbook(INPUT_EXCEL)
    ws = wb.active

    headers = {}
    for idx, cell in enumerate(ws[1]):
        if cell.value:
            headers[str(cell.value).strip()] = idx

    print("当前识别到的表头：", list(headers.keys()))

    contract_keys = ["合同编号", "合同号", "contractNo", "单号"]
    contract_index = next((headers[k] for k in contract_keys if k in headers), None)

    if contract_index is None:
        print("Excel 必须包含：合同编号")
        return []

    tasks = []

    for row in ws.iter_rows(min_row=2):
        contract_no = row[contract_index].value
        if contract_no:
            tasks.append({"contract_no": str(contract_no).strip()})

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

    try:
        wb.save(OUTPUT_EXCEL)
        print(f"结果已保存：{OUTPUT_EXCEL}")
    except PermissionError:
        alt_path = OUTPUT_EXCEL.replace(".xlsx", f"_{int(time.time())}.xlsx")
        wb.save(alt_path)
        print(f"output.xlsx 正在被打开，已另存为：{alt_path}")


def get_page(playwright):
    browser = playwright.chromium.connect_over_cdp(CDP_URL)
    context = browser.contexts[0]
    pages = [p for p in context.pages if not p.is_closed()]
    page = pages[0] if pages else context.new_page()
    return browser, context, page


def ensure_page(playwright, page):
    if page is None or page.is_closed():
        print("页面已关闭，重新连接 Chrome...")
        _, _, page = get_page(playwright)
    return page


def click_workbench_tab(page):
    tabs = page.locator(".ant-tabs-tab")
    count = tabs.count()

    print(f"检测到 {count} 个 tab")

    for i in range(count):
        tab = tabs.nth(i)
        text = tab.inner_text(timeout=3000).strip()
        print(f"TAB[{i}] => {text}")

        if text.startswith("电催工作台") and "详情" not in text:
            tab.click(force=True, timeout=10000)
            time.sleep(2)
            print("已切换到电催工作台")
            return

    raise Exception("没有找到电催工作台 TAB")


def search_contract(page, contract_no):
    print(f"开始搜索合同：{contract_no}")

    page.bring_to_front()
    time.sleep(1)

    # 直接找 placeholder
    contract_input = page.locator(
        'input.ant-input[placeholder*="批量搜索"]:not([disabled])'
    ).first

    contract_input.wait_for(timeout=30000)

    contract_input.scroll_into_view_if_needed(timeout=5000)

    contract_input.click(force=True, timeout=10000)

    # 清空
    contract_input.press("Control+A")
    contract_input.press("Backspace")

    # 输入
    contract_input.fill(contract_no, timeout=10000)

    print(f"已输入合同编号：{contract_no}")

    # 查询按钮
    query_btn = page.locator(
        'button.ant-btn-primary:has-text("查 询")'
    ).first

    query_btn.click(force=True, timeout=10000)

    print("已点击查询")

    # 等待表格出现
    row_selector = f'tr[data-row-key="{contract_no}"]'

    page.wait_for_selector(row_selector, timeout=30000)

    row = page.locator(row_selector).first

    # 点击合同编号
    contract_link = row.locator("a", has_text=contract_no).first

    contract_link.scroll_into_view_if_needed(timeout=5000)

    contract_link.click(force=True, timeout=10000)

    print("已点击合同编号进入详情")

    time.sleep(3)


def wait_detail_ready(page):
    page.wait_for_selector(".record", timeout=30000)
    page.wait_for_selector(".call-out", timeout=30000)

    page.wait_for_function(
        """
        () => {
            const text = document.body.innerText || ''
            return text.includes('合同编号') &&
                   text.includes('承租人') &&
                   text.includes('催记录入')
        }
        """,
        timeout=30000,
    )

    time.sleep(2)
    print("详情页数据已加载")


def get_current_status(page):
    try:
        status_select = page.locator("div.ant-select.current-status-value").first
        status_select.wait_for(state="visible", timeout=10000)
        text = status_select.inner_text(timeout=5000).strip()
        print("当前状态：", text)
        return text
    except Exception as e:
        print("获取状态失败：", e)
        return ""


def switch_status_to_idle(page):
    current_text = get_current_status(page)

    if "空闲" in current_text:
        print("当前已经是空闲状态")
        return

    print("当前不是空闲，开始切换为空闲...")

    status_select = page.locator("div.ant-select.current-status-value").first
    status_select.click(force=True, timeout=10000)

    time.sleep(1)

    dropdown = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden)").last
    dropdown.wait_for(state="visible", timeout=10000)

    idle_option = dropdown.locator(".ant-select-item-option").filter(has_text="空闲").last
    idle_option.scroll_into_view_if_needed(timeout=5000)
    idle_option.click(force=True, timeout=10000)

    page.wait_for_function(
        """
        () => {
            const el = document.querySelector('div.ant-select.current-status-value')
            return el && el.innerText.includes('空闲')
        }
        """,
        timeout=15000,
    )

    time.sleep(2)
    print("状态已切换为空闲")


def ensure_idle_status(page):
    for i in range(3):
        try:
            current = get_current_status(page)

            if "空闲" in current:
                print("状态正常：空闲")
                return

            switch_status_to_idle(page)

            current = get_current_status(page)

            if "空闲" in current:
                print("切换成功")
                return

        except Exception as e:
            print(f"第 {i + 1} 次切换失败：", e)

        time.sleep(2)

    raise Exception("无法切换为空闲状态")


def select_outbound_number(page):
    caller_select = page.locator("div.ant-select.dial-caller-select").first
    caller_select.wait_for(state="visible", timeout=20000)

    current_text = caller_select.inner_text(timeout=5000).strip()
    print("当前外显号码：", current_text)

    if current_text and "请选择" not in current_text:
        print("外显号码已存在")
        return

    caller_select.click(force=True, timeout=10000)
    time.sleep(1)

    dropdown = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden)").last
    dropdown.wait_for(state="visible", timeout=10000)

    option = dropdown.locator(".ant-select-item-option").filter(has_not_text="无数据").first
    option.click(force=True, timeout=10000)

    time.sleep(1)
    print("已选择外显号码")


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


def clean_phone(phone):
    return "".join(ch for ch in phone if ch.isdigit() or ch == "*")


def click_call_btn(page):
    page.locator(".call-out img[src*='contractMakeCall']").first.click(
        force=True,
        timeout=10000,
    )

    time.sleep(1)

    try:
        call_btn = page.locator("button.call-button:has-text('呼叫')").first
        if call_btn.is_visible(timeout=3000):
            call_btn.click(force=True, timeout=10000)
            print("已点击顶部呼叫")
    except Exception:
        pass

    print("已点击拨打")


def hang_up(page):
    print("等待 3 秒后挂断...")
    time.sleep(3)

    page.locator("button.call-button:has-text('挂断')").first.click(
        force=True,
        timeout=15000,
    )

    time.sleep(1)
    print("已挂断")


def get_visible_form_item_by_label(page, label_text):
    items = page.locator(f".record .ant-form-item:has(label[title='{label_text}'])")
    count = items.count()

    for i in range(count - 1, -1, -1):
        item = items.nth(i)
        try:
            box = item.bounding_box()
            if box and box["width"] > 0 and box["height"] > 0:
                return item
        except Exception:
            continue

    return items.last


def select_ant_option_by_label(page, label_text, option_text=None):
    form_item = get_visible_form_item_by_label(page, label_text)
    form_item.scroll_into_view_if_needed(timeout=8000)

    select_root = form_item.locator(".ant-select:not(.ant-select-disabled)").first
    select_root.click(force=True, timeout=10000)
    time.sleep(0.5)

    dropdown = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden)").last
    dropdown.wait_for(state="visible", timeout=10000)

    options = dropdown.locator(".ant-select-item-option").filter(has_not_text="无数据")

    if option_text:
        options.filter(has_text=option_text).last.click(force=True, timeout=10000)
    else:
        options.first.click(force=True, timeout=10000)

    time.sleep(0.5)
    print(f"已选择：{label_text} -> {option_text or '第一个'}")


def fill_input_by_label(page, label_text, value):
    form_item = get_visible_form_item_by_label(page, label_text)
    form_item.scroll_into_view_if_needed(timeout=8000)

    input_box = form_item.locator("input:not([disabled])").first
    input_box.fill(str(value), timeout=10000)

    time.sleep(0.3)
    print(f"已填写：{label_text} -> {value}")


def fill_contact_time(page):
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    page.evaluate(
        """
        (value) => {
            const picker = document.querySelector('#contactTime')
            if (!picker) return

            const input = picker.querySelector('input')
            if (!input) return

            input.removeAttribute('disabled')
            input.value = value

            input.dispatchEvent(new Event('input', { bubbles: true }))
            input.dispatchEvent(new Event('change', { bubbles: true }))
            input.dispatchEvent(new Event('blur', { bubbles: true }))
        }
        """,
        now_text,
    )

    time.sleep(0.5)
    print("已填写联络时间：", now_text)


def wait_call_record_form_ready(page):
    page.wait_for_selector("#riskType", timeout=30000)
    page.wait_for_selector("#contactResult", timeout=30000)
    page.wait_for_selector("#collectType", timeout=30000)
    page.wait_for_selector("#touch", timeout=30000)

    time.sleep(1)
    print("催记录入表单已就绪")


def fill_collection_form(page, phone):
    wait_call_record_form_ready(page)

    try:
        select_ant_option_by_label(page, "催收形式", "外呼")
    except Exception as e:
        print("催收形式选择失败，继续：", e)

    phone_value = clean_phone(phone)

    if phone_value:
        try:
            fill_input_by_label(page, "电话", phone_value)
        except Exception as e:
            print("电话填写失败，继续：", e)

    select_ant_option_by_label(page, "风险分类", "失联")

    try:
        select_ant_option_by_label(page, "是否触达", "未触达")
    except Exception as e:
        print("是否触达选择失败，继续：", e)

    try:
        fill_contact_time(page)
    except Exception as e:
        print("联络时间填写失败，继续：", e)

    select_ant_option_by_label(page, "联络结果")


def submit_form(page):
    record = page.locator(".record").last
    submit_btn = record.locator("button.ant-btn-primary:has-text('提 交')").last

    submit_btn.scroll_into_view_if_needed(timeout=8000)
    submit_btn.click(force=True, timeout=10000)

    time.sleep(2)
    print("已提交")


def process_case(page, task):
    contract_no = task["contract_no"]

    print(f"开始处理：{contract_no}")

    click_workbench_tab(page)

    search_contract(page, contract_no)

    wait_detail_ready(page)

    ensure_idle_status(page)

    select_outbound_number(page)

    ensure_idle_status(page)

    name = get_name_from_page(page)
    phone = get_real_phone(page)

    click_call_btn(page)

    hang_up(page)

    fill_collection_form(page, phone)

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
    try:
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
            print("请确认 Chrome 已登录系统，并停留在电催工作台页面")
            print("确认后按回车继续")
            print("=" * 50)
            input()

            for index, task in enumerate(tasks, start=1):
                print(f"\n[{index}/{len(tasks)}]")

                try:
                    page = ensure_page(p, page)
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

                    try:
                        page = ensure_page(p, page)
                    except Exception:
                        pass

                results.append(result)
                save_results(results)

            print("\n全部完成")
            print(f"结果文件：{OUTPUT_EXCEL}")
            print("Chrome 不会关闭，exe 窗口也不会自动关闭")

    except Exception:
        print("程序发生未捕获异常：")
        print(traceback.format_exc())

    finally:
        pause_exit()


if __name__ == "__main__":
    main()
