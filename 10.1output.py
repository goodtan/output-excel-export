import requests
import urllib3
import json
import pandas as pd
import os

# 禁用不安全请求警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def parse_cookie_string(cookie_str):
    """解析浏览器复制的 Raw Cookie 字符串为字典"""
    cookies = {}
    if not cookie_str:
        return cookies
    for part in cookie_str.split(";"):
        part = part.strip()
        if not part:
            continue
        if "=" in part:
            k, v = part.split("=", 1)
            cookies[k.strip()] = v.strip()
        else:
            cookies[part] = ""
    return cookies

def read_file_content(filename):
    """读取当前目录下文件的内容，如果文件不存在返回空字符串"""
    if not os.path.exists(filename):
        print(f"⚠️  警告: 未找到文件 [{filename}]，将跳过该参数。")
        return ""
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            # 去掉可能存在的换行符或多余空格
            return content.replace('\n', '').replace('\r', '').strip()
    except Exception as e:
        print(f"❌ 读取文件 [{filename}] 出错: {e}")
        return ""

def main():
    print("========================================================")
    print("正在初始化... 准备从本地文件读取认证信息")
    print("目标接口: https://120.55.38.129:9998/.../customerList")
    print("========================================================")

    # 1. 从文件读取认证信息
    # 文件名定义
    file_authorization = "authorization.txt"
    file_blade_auth = "blade_auth.txt"
    file_cookies = "cookies.txt"

    print(f"正在读取 {file_authorization} ...")
    authorization_val = read_file_content(file_authorization)

    print(f"正在读取 {file_blade_auth} ...")
    blade_auth_val = read_file_content(file_blade_auth)

    print(f"正在读取 {file_cookies} ...")
    cookie_str = read_file_content(file_cookies)

    # 简单检查
    if not cookie_str:
        print("\n❌ 错误: cookies.txt 内容为空或文件不存在，无法进行请求。")
        input("按回车键退出...")
        return

    # 2. 手动输入页码（因为每次导出范围可能不同）
    print("\n--------------------------------------------------------")
    start_page_input = input("请输入起始页码 (默认 1): ").strip()
    start_page = int(start_page_input) if start_page_input.isdigit() else 1

    end_page_input = input(f"请输入结束页码 (默认 {start_page}): ").strip()
    end_page = int(end_page_input) if end_page_input.isdigit() else start_page
    print("--------------------------------------------------------\n")

    # 3. 构造请求头
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Connection": "keep-alive",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://120.55.38.129:9998",
        "Referer": "https://120.55.38.129:9998/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    }

    if authorization_val:
        headers["Authorization"] = authorization_val
    if blade_auth_val:
        headers["blade-Auth"] = blade_auth_val

    # 4. 构造 Cookies
    cookies = parse_cookie_string(cookie_str)

    url = "https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList"
    all_records = []

    # 5. 循环请求
    for page in range(start_page, end_page + 1):
        payload = {
            "current": page,
            "size": 20,
            "customerId": "",
            "customerData": "",
            "borrowerName": "",
            "idNo": "",
            "projectTypeList": [],
            "originalCreditor": "",
            "returnStatus": "",
            "flagStatus": "",
            "contractNo": "",
            "sumSurplusPrincipal": "",
            "sumSurplusLoan": "",
            "phone": "",
            "tenantId": "831444",
            "deptCompanyId": "",
            "salesmanName": "",
            "unusedDays": "",
            "commissionDays": "",
            "paySchedule": "",
            "surplusLoanLeft": "",
            "surplusLoanRight": "",
            "surplusPrincipalLeft": "",
            "surplusPrincipalRight": "",
            "overdueDaysLeft": "",
            "overdueDaysRight": "",
            "sumSurplusLoanLeft": "",
            "sumSurplusLoanRight": "",
            "sumSurplusPrincipalLeft": "",
            "sumSurplusPrincipalRight": "",
            "followTime": "",
            "payedAmount": "",
            "commissionDaysLeft": "",
            "commissionDaysRight": "",
            "payScheduleLeft": "",
            "payScheduleRight": ""
        }

        print(f"正在请求第 {page} 页 ...", end="")
        try:
            resp = requests.post(url, headers=headers, cookies=cookies, json=payload,
                                 verify=False, timeout=30)
            
            if resp.status_code == 200:
                data = resp.json()
                # 检查业务状态码 (有些系统 HTTP 200 但 code!=200 代表token过期)
                if data.get("code") == 401:
                    print(f" 失败 -> Token 已过期，请更新 cookies.txt 或 header 文件")
                    break
                
                records = data.get("data", {}).get("records", [])
                count = len(records)
                print(f" 成功 (获取到 {count} 条数据)")
                all_records.extend(records)
                
                # 如果获取的数据少于 size(20)，说明是最后一页了，可以提前结束
                if count < 20:
                    print("已到达最后一页，停止翻页。")
                    break
            else:
                print(f" 失败 (HTTP {resp.status_code})")

        except Exception as e:
            print(f" 异常: {e}")

    # 6. 导出数据
    if not all_records:
        print("\n没有获取到任何数据，未生成 Excel。")
    else:
        output_file = "customerList.xlsx"
        print(f"\n正在写入 Excel: {output_file} ...")
        try:
            df = pd.DataFrame(all_records)
            df.to_excel(output_file, index=False)
            print(f"✅ 成功! 共导出 {len(all_records)} 条记录。")
        except Exception as e:
            print(f"❌ 写入 Excel 失败: {e}")
            print("请检查文件是否被其他程序（如 WPS/Excel）占用。")

    input("\n任务完成，按回车键退出...")

if __name__ == "__main__":
    main()
