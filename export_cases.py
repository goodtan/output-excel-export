import os
import sys
import time
import traceback
from datetime import datetime

from openpyxl import load_workbook, Workbook
from playwright.sync_api import sync_playwright


INPUT_EXCEL_NAME = "input.xlsx"
OUTPUT_EXCEL_NAME = "output.xlsx"

USE_EXISTING_CHROME = True
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
    url_keys = ["详情URL", "详情页URL", "URL", "url", "链接", "详情页链接"]

    contract_index = next((headers[k] for k in contract_keys if k in headers), None)
    url_index = next((headers[k] for k in url_keys if k in headers), None)

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
    try:
        if page is None or page.is_closed():
            print("页面已关闭，正在重新连接 Chrome...")
            _, _, page = get_page(playwright)
        return page
    except Exception:
        print("重新连接 Chrome 失败，请确认 Chrome 远程调试窗口还开着")
        raise


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


def sign_in(page):
    try:
        btn = page.locator("button.ant-switch").filter(has_text="签入").first
        btn.wait_for(state="visible", timeout=10000)

        cls = btn.get_attribute("class") or ""
        print("签入按钮状态：", cls)

        if "ant-switch-checked" in cls:
            print("当前已经签入")
            return

        btn.click(force=True, timeout=10000)

        page.wait_for_function(
            """
            () => {
                const btns = [...document.querySelectorAll('button.ant-switch')]
                const btn = btns.find(b => b.innerText.includes('签入'))
                return btn && btn.className.includes('ant-switch-checked')
            }
            """,
            timeout=15000,
        )

        time.sleep(2)
        print("签入成功")

    except Exception as e:
        print("签入失败或已经签入，跳过：", e)


def change_current_status(page):
    status_select = page.locator("div.ant-select.current-status-value").first
    status_select.wait_for(state="visible", timeout=20000)

    current_text = status_select.inner_text(timeout=5000)
    print("当前状态：", current_text)

    if "空闲" in current_text:
        print("当前已是空闲")
        return

    status_select.click(force=True, timeout=10000)
    time.sleep(1)

    dropdown = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden)").last
    dropdown.wait_for(state="visible", timeout=10000)

    dropdown.locator(".ant-select-item-option").filter(has_text="空闲").last.click(
        force=True,
        timeout=10000,
    )

    time.sleep(2)

    page.wait_for_function(
        """
        () => {
            const el = document.querySelector('div.ant-select.current-status-value')
            return el && el.innerText.includes('空闲')
        }
        """,
        timeout=15000,
    )

    print("已切换状态为空闲")


def ensure_idle(page):
    for _ in range(3):
        try:
            change_current_status(page)
            return
        except Exception as e:
            print("状态不是空闲，重试切换：", e)
            time.sleep(2)

    raise Exception("状态无法切换为空闲")


def select_outbound_number(page):
    caller_select = page.locator("div.ant-select.dial-caller-select").first
    caller_select.wait_for(state="visible", timeout=20000)

    current_text = caller_select.inner_text(timeout=5000).strip()
    print("当前外显号码：", current_text)

    if current_text and "请选择" not in current_text:
        print("外显号码已存在")
        ensure_idle(page)
        return

    caller_select.click(force=True, timeout=10000)
    time.sleep(1)

    dropdown = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden)").last
    dropdown.wait_for(state="visible", timeout=10000)

    option = dropdown.locator(".ant-select-item-option").filter(has_not_text="无数据").first
    option.scroll_into_view_if_needed(timeout=5000)
    option.click(force=True, timeout=10000)

    time.sleep(1)
    print("已选择外显号码")

    ensure_idle(page)


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
    ensure_idle(page)

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
    page.wait_for_selector("#riskType", timeout=30000)
    page.wait_for_selector("#contactResult", timeout=30000)
    page.wait_for_selector("#collectType", timeout=30000)
    page.wait_for_selector("#touch", timeout=30000)

    time.sleep(1)
    print("催记录入表单已就绪")


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


def fill_collection_form(page, phone):
    wait_call_record_form_ready(page)

    # 催收形式：外呼
    try:
        select_ant_option_by_label(page, "催收形式", "外呼")
    except Exception as e:
        print("催收形式选择失败，继续：", e)

    # 电话
    phone_value = clean_phone(phone)
    if phone_value:
        try:
            fill_input_by_label(page, "电话", phone_value)
        except Exception as e:
            print("电话填写失败，继续：", e)

    # 风险分类：失联
    select_ant_option_by_label(page, "风险分类", "失联")

    # 是否触达：未触达
    try:
        select_ant_option_by_label(page, "是否触达", "未触达")
    except Exception as e:
        print("是否触达选择失败，继续：", e)

    # 联络时间
    try:
        fill_contact_time(page)
    except Exception as e:
        print("联络时间填写失败，继续：", e)

    # 联络结果：随便选一个
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
    detail_url = task["detail_url"]

    print(f"开始处理：{contract_no}")

    page.goto(detail_url, wait_until="domcontentloaded")

    wait_detail_ready(page)

    sign_in(page)

    ensure_idle(page)

    select_outbound_number(page)

    ensure_idle(page)

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
            print("请确认 Chrome 已登录系统")
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
