# # request_customer_list_to_excel.py
# import requests
# import urllib3
# import pandas as pd
# import os
# import sys
# import datetime

# urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# def parse_cookie_string(cookie_str):
#     cookies = {}
#     if not cookie_str:
#         return cookies
#     for part in cookie_str.split(";"):
#         part = part.strip()
#         if not part:
#             continue
#         if "=" in part:
#             k, v = part.split("=", 1)
#             cookies[k.strip()] = v.strip()
#         else:
#             cookies[part] = ""
#     return cookies


# def input_nonempty(prompt):
#     return input(prompt).strip()


# def main():
#     print("向 https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList 发起请求并导出 Excel")

#     # 输入认证信息
#     authorization = input_nonempty("Authorization header (例如：Basic ...)：")
#     blade_auth = input_nonempty("blade-Auth header (例如：bearer ...)：")
#     saber_access_token = input_nonempty("saber-access-token cookie 值：")
#     saber_refresh_token = input_nonempty("saber-refresh-token cookie 值：")
#     extra_cookies = input("额外 cookie（可选，例如 JG_...=value; other=val）：").strip()

#     start_page = int(input("请输入起始页码：").strip() or "1")
#     end_page = int(input("请输入结束页码：").strip() or str(start_page))

#     headers = {
#         "Accept": "application/json, text/plain, */*",
#         "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
#         "Connection": "keep-alive",
#         "Content-Type": "application/json;charset=UTF-8",
#         "Origin": "https://120.55.38.129:9998",
#         "Referer": "https://120.55.38.129:9998/",
#         "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36 Edg/140.0.0.0",
#     }
#     if authorization:
#         headers["Authorization"] = authorization
#     if blade_auth:
#         headers["blade-Auth"] = blade_auth

#     cookies = {}
#     if saber_access_token:
#         cookies["saber-access-token"] = saber_access_token
#     if saber_refresh_token:
#         cookies["saber-refresh-token"] = saber_refresh_token
#     if extra_cookies:
#         cookies.update(parse_cookie_string(extra_cookies))

#     url = "https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList"

#     all_records = []

#     for page in range(start_page, end_page + 1):
#         payload = {
#             "current": page,
#             "size": 20,
#             "customerId": "",
#             "customerData": "",
#             "borrowerName": "",
#             "idNo": "",
#             "projectTypeList": [],
#             "originalCreditor": "",
#             "returnStatus": "",
#             "flagStatus": "",
#             "contractNo": "",
#             "sumSurplusPrincipal": "",
#             "sumSurplusLoan": "",
#             "phone": "",
#             "tenantId": "831444",
#             "deptCompanyId": "",
#             "salesmanName": "",
#             "unusedDays": "",
#             "commissionDays": "",
#             "paySchedule": "",
#             "surplusLoanLeft": "",
#             "surplusLoanRight": "",
#             "surplusPrincipalLeft": "",
#             "surplusPrincipalRight": "",
#             "overdueDaysLeft": "",
#             "overdueDaysRight": "",
#             "sumSurplusLoanLeft": "",
#             "sumSurplusLoanRight": "",
#             "sumSurplusPrincipalLeft": "",
#             "sumSurplusPrincipalRight": "",
#             "followTime": "",
#             "payedAmount": "",
#             "commissionDaysLeft": "",
#             "commissionDaysRight": "",
#             "payScheduleLeft": "",
#             "payScheduleRight": ""
#         }

#         print(f"请求第 {page} 页 ...")
#         try:
#             resp = requests.post(url, headers=headers, cookies=cookies, json=payload,
#                                  verify=False, timeout=30)
#             if resp.status_code != 200:
#                 print(f"第 {page} 页请求失败，HTTP {resp.status_code}")
#                 continue
#             data = resp.json()
#             records = data.get("data", {}).get("records", [])
#             print(f"第 {page} 页获取到 {len(records)} 条记录")
#             all_records.extend(records)
#         except Exception as e:
#             print(f"第 {page} 页请求异常: {e}")

#     if not all_records:
#         print("没有获取到任何数据，结束。")
#         input("按回车键退出...")
#         return

#     # 获取输出目录（脚本或 exe 所在目录）
#     if getattr(sys, 'frozen', False):
#         base_dir = os.path.dirname(sys.executable)
#     else:
#         base_dir = os.path.dirname(os.path.abspath(__file__))

#     # 带时间戳的文件名
#     timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
#     output_file = os.path.join(base_dir, f"customerList_{timestamp}.xlsx")

#     df = pd.DataFrame(all_records)
#     df.to_excel(output_file, index=False)
#     print(f"成功导出 {len(all_records)} 条记录到 {output_file}")

#     input("任务完成，按回车键退出...")


# if __name__ == "__main__":
#     main()




# # request_customer_list_to_excel_fixed.py
# import requests
# import urllib3
# import pandas as pd
# import os
# import sys
# import datetime
# import json

# urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# def main():
#     print("向 https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList 发起请求并导出 Excel")

#     # === 输入认证信息 ===
#     authorization = input("Authorization (例如 Basic ...): ").strip()
#     blade_auth = input("blade-Auth (例如 bearer ...): ").strip()
#     saber_access_token = input("saber-access-token cookie 值: ").strip()
#     saber_refresh_token = input("saber-refresh-token cookie 值: ").strip()
#     jg_cookie = input("JG_ 开头 cookie 值 (例如 JG_xxx_PV=...|...): ").strip()

#     start_page = int(input("起始页码: ").strip() or "1")
#     end_page = int(input("结束页码: ").strip() or str(start_page))

#     url = "https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList"

#     # === Headers，完全模拟浏览器 ===
#     headers = {
#         "Accept": "application/json, text/plain, */*",
#         "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
#         "Authorization": authorization,
#         "blade-Auth": blade_auth,
#         "Connection": "keep-alive",
#         "Content-Type": "application/json;charset=UTF-8",
#         "Origin": "https://120.55.38.129:9998",
#         "Referer": "https://120.55.38.129:9998/",
#         "Sec-Fetch-Dest": "empty",
#         "Sec-Fetch-Mode": "cors",
#         "Sec-Fetch-Site": "same-origin",
#         "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36 Edg/141.0.0.0",
#         "sec-ch-ua": '"Microsoft Edge";v="141", "Not?A_Brand";v="8", "Chromium";v="141"',
#         "sec-ch-ua-mobile": "?0",
#         "sec-ch-ua-platform": '"Windows"',
#     }

#     # === Cookies，与 cURL 一致 ===
#     cookies = {
#         "saber-access-token": saber_access_token,
#         "saber-refresh-token": saber_refresh_token,
#     }
#     if jg_cookie:
#         # 自动判断 cookie 名（例如 "JG_xxx_PV"）
#         if "=" in jg_cookie:
#             name, value = jg_cookie.split("=", 1)
#             cookies[name.strip()] = value.strip()

#     all_records = []

#     for page in range(start_page, end_page + 1):
#         payload = {
#             "current": page,
#             "size": 20,
#             "province": "",
#             "city": "",
#             "area": "",
#             "customerId": "",
#             "customerData": "",
#             "borrowerName": "",
#             "idNo": "",
#             "projectTypeList": [],
#             "originalCreditor": "",
#             "returnStatus": "",
#             "flagStatus": "",
#             "contractNo": "",
#             "sumSurplusPrincipal": "",
#             "sumSurplusLoan": "",
#             "phone": "",
#             "deptCompanyId": "",
#             "salesmanName": "",
#             "unusedDays": "",
#             "commissionDays": "",
#             "paySchedule": "",
#             "tenantId": "831444",
#             "surplusLoanLeft": "",
#             "surplusLoanRight": "",
#             "surplusPrincipalLeft": "",
#             "surplusPrincipalRight": "",
#             "identyStatus": "",
#             "payScheduleLeft": "",
#             "payScheduleRight": "",
#         }

#         print(f"请求第 {page} 页 ...")
#         try:
#             resp = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False, timeout=30)
#             if resp.status_code == 401:
#                 print(f"⚠️ 第 {page} 页返回 401（认证失败）")
#                 print(resp.text)
#                 continue
#             if resp.status_code != 200:
#                 print(f"⚠️ 第 {page} 页 HTTP {resp.status_code}")
#                 continue

#             data = resp.json()
#             records = data.get("data", {}).get("records", [])
#             print(f"✅ 第 {page} 页获取 {len(records)} 条记录")
#             all_records.extend(records)
#         except Exception as e:
#             print(f"❌ 第 {page} 页请求异常: {e}")

#     if not all_records:
#         print("❗未获取到任何数据。")
#         input("按回车键退出...")
#         return

#     # === 保存 Excel ===
#     base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
#     timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
#     output_file = os.path.join(base_dir, f"customerList_{timestamp}.xlsx")

#     pd.DataFrame(all_records).to_excel(output_file, index=False)
#     print(f"✅ 成功导出 {len(all_records)} 条记录到 {output_file}")
#     input("任务完成，按回车键退出...")

# if __name__ == "__main__":
#     main()


import re
import json
import requests

def parse_curl(curl_str: str):
    """解析 curl 命令，提取 URL、headers、cookies、data"""
    url_match = re.search(r"curl\s+[\"']?([^\"'\s]+)", curl_str)
    url = url_match.group(1) if url_match else None

    headers = {}
    for h in re.findall(r"-H\s+[\"']([^\"']+)[\"']", curl_str):
        if ":" in h:
            key, value = h.split(":", 1)
            headers[key.strip()] = value.strip()

    # cookies
    cookie_match = re.search(r"-b\s+[\"']([^\"']+)[\"']", curl_str)
    cookies = {}
    if cookie_match:
        for c in cookie_match.group(1).split(";"):
            if "=" in c:
                k, v = c.split("=", 1)
                cookies[k.strip()] = v.strip()

    # data
    data_match = re.search(r"--data-raw\s+[\"'](.+?)[\"'](\s|$)", curl_str)
    data = data_match.group(1) if data_match else None
    if data:
        try:
            data = json.loads(data)
        except Exception:
            pass

    return url, headers, cookies, data


def run_request_from_curl(curl_str: str):
    """执行请求并打印结果"""
    url, headers, cookies, data = parse_curl(curl_str)
    if not url:
        print("❌ 未找到 URL，请检查 curl 格式。")
        return

    print(f"🌐 请求 URL: {url}")
    try:
        resp = requests.post(url, headers=headers, cookies=cookies, json=data, verify=False, timeout=15)
        if resp.status_code == 200:
            print("✅ 请求成功！返回内容：")
            print(json.dumps(resp.json(), ensure_ascii=False, indent=2))
        else:
            print(f"⚠️ 请求失败，状态码: {resp.status_code}")
            print("返回内容:")
            print(resp.text)
    except Exception as e:
        print(f"❌ 请求出错: {e}")


if __name__ == "__main__":
    print("请输入完整 curl 命令（按回车结束输入）：")
    curl_input = []
    while True:
        try:
            line = input()
            if not line:
                break
            curl_input.append(line)
        except EOFError:
            break

    curl_str = " ".join(curl_input)
    run_request_from_curl(curl_str)


