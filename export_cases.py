import os
import sys
import time
import traceback
from datetime import datetime

from openpyxl import load_workbook, Workbook
from playwright.sync_api import sync_playwright


INPUT_EXCEL_NAME = "input.xlsx"
OUTPUT_EXCEL_NAME = f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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

    wb.save(OUTPUT_EXCEL)
    print(f"结果已保存：{OUTPUT_EXCEL}")


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

    contract_input = page.locator(
        'input.ant-input[placeholder*="批量搜索"]:not([disabled])'
    ).first

    contract_input.wait_for(timeout=30000)
    contract_input.scroll_into_view_if_needed(timeout=5000)
    contract_input.click(force=True, timeout=10000)
    contract_input.press("Control+A")
    contract_input.press("Backspace")
    contract_input.fill(contract_no, timeout=10000)

    print(f"已输入合同编号：{contract_no}")

    query_btn = page.locator('button.ant-btn-primary:has-text("查 询")').first
    query_btn.click(force=True, timeout=10000)
    print("已点击查询")

    row_selector = f'tr[data-row-key="{contract_no}"]'
    page.wait_for_selector(row_selector, timeout=30000)

    row = page.locator(row_selector).first
    row.scroll_into_view_if_needed(timeout=5000)
    time.sleep(1)

    contract_link = row.locator("a", has_text=contract_no).first

    try:
        contract_link.click(force=True, timeout=8000)
        print("已点击合同编号")
    except Exception as e:
        print("点击 a 标签失败，改为双击整行：", e)
        row.dblclick(force=True, timeout=8000)

    time.sleep(5)
    print("已尝试进入详情")


def wait_detail_ready(page):
    page.wait_for_selector(".record", timeout=30000)
    page.wait_for_selector(".call-out", timeout=30000)

    page.wait_for_function(
        """
        () => {
            const text = document.body.innerText || ''
            return text.includes('承租人') && text.includes('催记录入')
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
        call_out = page.locator(".call-out").first
        call_out.wait_for(timeout=10000)

        try:
            call_out.locator(".toggle-des").first.click(force=True, timeout=5000)
            time.sleep(1)
        except Exception as e:
            print("点击手机号展示按钮失败，继续尝试读取：", e)

        spans = call_out.locator("span[title]")
        count = spans.count()

        for i in range(count):
            text = spans.nth(i).inner_text(timeout=3000).strip()
            if text.isdigit() and len(text) == 11:
                print("获取到手机号：", text)
                return text

        text = call_out.inner_text(timeout=5000).replace("\n", " ").strip()
        print("兜底手机号文本：", text)
        return text

    except Exception as e:
        print("获取手机号失败：", e)
        return ""


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


def wait_call_record_form_ready(page):
    form = page.locator(".add-collection-record").last
    form.wait_for(timeout=30000)

    risk = form.locator("#riskType").last
    risk.wait_for(timeout=30000)

    print("催记录入表单已就绪")


def get_form_item_by_label(page, label_text):
    form = page.locator(".add-collection-record").last
    items = form.locator(f'.ant-form-item:has(label[title="{label_text}"])')
    count = items.count()

    for i in range(count - 1, -1, -1):
        item = items.nth(i)
        try:
            box = item.bounding_box()
            if box and box["width"] > 0 and box["height"] > 0:
                return item
        except Exception:
            pass

    return items.last


def select_ant_option_by_label(page, label_text, option_text=None):
    item = get_form_item_by_label(page, label_text)
    item.scroll_into_view_if_needed(timeout=8000)

    select_root = item.locator(".ant-select:not(.ant-select-disabled)").last
    select_root.click(force=True, timeout=10000)
    time.sleep(1)

    dropdown = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden)").last
    dropdown.wait_for(timeout=10000)

    if not option_text:
        option = dropdown.locator(".ant-select-item-option").filter(has_not_text="无数据").first
        option.click(force=True, timeout=10000)
        print(f"已选择：{label_text} -> 第一个")
        return

    for _ in range(30):
        option = dropdown.locator(
            f'.ant-select-item-option[title="{option_text}"], '
            f'.ant-select-item-option[label="{option_text}"]'
        ).last

        if option.count() > 0:
            option.scroll_into_view_if_needed(timeout=3000)
            option.click(force=True, timeout=10000)
            time.sleep(0.5)
            print(f"已选择：{label_text} -> {option_text}")
            return

        holder = dropdown.locator(".rc-virtual-list-holder").first
        holder.evaluate("(el) => { el.scrollTop = el.scrollTop + 220 }")
        time.sleep(0.25)

    raise Exception(f"未找到选项：{label_text} -> {option_text}")


def fill_collection_form(page):
    wait_call_record_form_ready(page)

    try:
        select_ant_option_by_label(page, "风险分类", "失联")
    except Exception as e:
        print("风险分类选择失败：", e)

    try:
        select_ant_option_by_label(page, "联络结果", "无法接通")
    except Exception as e:
        print("联络结果选择失败：", e)

    print("催记录入填写完成")


def submit_form(page):
    record = page.locator(".add-collection-record").last
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

    fill_collection_form(page)
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
