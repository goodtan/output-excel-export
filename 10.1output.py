# request_customer_list.py
import requests
import urllib3
import json

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def parse_cookie_string(cookie_str):
    """
    把形如 "a=1; b=2" 的 cookie 字符串解析成 dict
    """
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
    v = input(prompt).strip()
    return v

def main():
    print("向 https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList 发起 POST 请求")
    print("按回车留空表示不填写该项（某些项若为空可能导致请求被拒绝）")
    authorization = input_nonempty("Authorization header (例如：Basic ...)：")
    blade_auth = input_nonempty("blade-Auth header (例如：bearer ...)：")
    saber_access_token = input_nonempty("saber-access-token cookie 值：")
    saber_refresh_token = input_nonempty("saber-refresh-token cookie 值：")
    extra_cookies = input("额外 cookie（可选，例如 JG_...=value; other=val），留空则无：").strip()

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://120.55.38.129:9998",
        "Referer": "https://120.55.38.129:9998/",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36 Edg/140.0.0.0",
        "sec-ch-ua": "\"Chromium\";v=\"140\", \"Not=A?Brand\";v=\"24\", \"Microsoft Edge\";v=\"140\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Windows\""
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

    payload = {
        "current": 1,
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

    url = "https://120.55.38.129:9998/api/blade-system/baseCaseNew/customerList"

    try:
        print("\n发起请求（verify=False，对应 curl --insecure）...")
        resp = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False, timeout=30)
    except Exception as e:
        print("请求失败：", e)
        return

    print("\nHTTP 状态码：", resp.status_code)
    # 尝试解析 JSON
    text = resp.text
    try:
        parsed = resp.json()
        print("响应（JSON，格式化）：")
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
    except Exception:
        print("响应（非 JSON 或解析失败），原始文本：")
        print(text)

if __name__ == "__main__":
    main()
