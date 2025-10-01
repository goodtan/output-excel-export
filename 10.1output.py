# request_customer_list_to_excel.py
import requests
import urllib3
import json
import pandas as pd

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def parse_cookie_string(cookie_str):
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

def input_nonempty(prompt):
    return input(prompt).strip()

def main():
    print("向 https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList 发起请求并导出 Excel")

    # 输入认证信息
    authorization = input_nonempty("Authorization header (例如：Basic ...)：")
    blade_auth = input_nonempty("blade-Auth header (例如：bearer ...)：")
    saber_access_token = input_nonempty("saber-access-token cookie 值：")
    saber_refresh_token = input_nonempty("saber-refresh-token cookie 值：")
    extra_cookies = input("额外 cookie（可选，例如 JG_...=value; other=val）：").strip()

    start_page = int(input("请输入起始页码：").strip() or "1")
    end_page = int(input("请输入结束页码：").strip() or str(start_page))

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://120.55.38.129:9998",
        "Referer": "https://120.55.38.129:9998/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36 Edg/140.0.0.0",
    }
    if authorization:
        headers["Authorization"] = authorization
    if blade_auth:
        headers["blade-Auth"] = blade_auth

    cookies = {}
    if saber_access_token:
        cookies["saber-access-token"] = saber_access_token
    if saber_refresh_token:
        cookies["saber-refresh-token"] = saber_refresh_token
    if extra_cookies:
        cookies.update(parse_cookie_string(extra_cookies))

    url = "https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList"

    all_records = []

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

        print(f"请求第 {page} 页 ...")
        try:
            resp = requests.post(url, headers=headers, cookies=cookies, json=payload,
                                 verify=False, timeout=30)
            if resp.status_code != 200:
                print(f"第 {page} 页请求失败，HTTP {resp.status_code}")
                continue
            data = resp.json()
            records = data.get("data", {}).get("records", [])
            print(f"第 {page} 页获取到 {len(records)} 条记录")
            all_records.extend(records)
        except Exception as e:
            print(f"第 {page} 页请求异常: {e}")

    if not all_records:
        print("没有获取到任何数据，结束。")
    else:
        # 导出 Excel
        df = pd.DataFrame(all_records)
        output_file = "customerList.xlsx"
        df.to_excel(output_file, index=False)
        print(f"成功导出 {len(all_records)} 条记录到 {output_file}")

    # 👇 这里加上暂停，不让窗口自动退出
    input("任务完成，按回车键退出...")

if __name__ == "__main__":
    main()
