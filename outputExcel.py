# import requests

# url = "https://gateway.fangnuokeji.com/caseCenter/case/allot/orgAllotCaseList"

# headers = {
#     "accept": "application/json, text/plain, */*",
#     "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
#     "authorization": "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzeXNfc2VydmljZV91c2VyX25hbWUiOiLlrovkuJbpvpkiLCJzeXNfcHJvZHVjdF9pZCI6IjYwZTkxOTk2NGQwYWE1YWY3YzBkOGEwNjU3MTc5YzY2YWYyZTQzNGIiLCJzeXNfdXNlcl9tb2JpbGVfcGhvbmUiOiIxMzU2OTQ3NzA0NiIsInVzZXJfbmFtZSI6IuWui-S4lum-mSIsInNjb3BlIjpbImFsbCJdLCJzeXNfdXNlcl9uYW1lIjoi5a6L5LiW6b6ZIiwic3lzX3NlcnZpY2VfdXNlcl9pZCI6NjA5MjYsImV4cCI6MTc1MzU1NzczNSwianRpIjoiMGEyN2FhYTctNTJjNS00ZGUwLTgzZGUtZjRlNDQ5YjdkOWIxIiwiZ2xvYmFsX3VzZXJfdXVpZCI6Ijk2OGVmOWExNjM4NWNmZjA1YWVjN2Q0YjQ2ZjAzNWRlIiwiYXV0aG9yaXRpZXMiOlsiYWRtaW4mSklFU1VBTkRBTiJdLCJjbGllbnRfaWQiOiLkuJzlsrjmmbrog73lpITnva7ns7vnu58tQ1BFIn0.NvBFB7gH-PDn0BdFZhGz8pk23FAj1AJYF1dkb2Lfp-q3GYnNsvUGvtYbNDjJhf9Ap20RzMnCC11LmT8B9dBe1DkcPPgzQMa9Q4pJLlBYTaLiH0fmFH8HIo5vAHPbt6bRs3u3uqpiky3ltd0FVLXF0wQL3SH4Ojc_Dx8P7IGX217mYAGZHUfaod6MmLKdtQVvFW0sJvmwUM_zZ9XoLWuXGqVroPjjfsQ1bOssgKV_nqcZ6yL89FwdKHmIarpb_c7jAVHr51R18IvEls0NvSpD8shXSPf15k5_XdM2q1VA0FWmxa_Dodl0WriTFMMhqA1SrZY0q5yo2OTqWfnFRfACRA",  # 建议替换为最新 token
#     "content-type": "application/json;charset=UTF-8",
#     "origin": "https://disposal.fangnuokeji.com",
#     "priority": "u=1, i",
#     "referer": "https://disposal.fangnuokeji.com/",
#     "sec-ch-ua": "\"Not)A;Brand\";v=\"8\", \"Chromium\";v=\"138\", \"Microsoft Edge\";v=\"138\"",
#     "sec-ch-ua-mobile": "?0",
#     "sec-ch-ua-platform": "\"Windows\"",
#     "sec-fetch-dest": "empty",
#     "sec-fetch-mode": "cors",
#     "sec-fetch-site": "same-site",
#     "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36 Edg/138.0.0.0"
# }

# payload = {
#     "caseNo": None,
#     "productId": None,
#     "entrustBatchId": None,
#     "batchTypeId": None,
#     "userName": None,
#     "idno": None,
#     "userPhone": None,
#     "followStatusId": None,
#     "refundStatus": None,
#     "caseStatus": None,
#     "isRetain": None,
#     "retainStagingPlanStatus": None,
#     "stagingPlanStatus": None,
#     "isTagAlter": None,
#     "isFollow": None,
#     "entrustContactResultIdList": None,
#     "color": None,
#     "cpeId": None,
#     "startallotTime": None,
#     "endallotTime": None,
#     "retainEndTimeStart": None,
#     "retainEndTimeEnd": None,
#     "entrustAmountMax": None,
#     "entrustAmountMin": None,
#     "regAddrProvince": None,
#     "regAddrCity": None,
#     "regAddrArea": None,
#     "investorName": None,
#     "orgTagTempName": None,
#     "caseLevelDesc": None,
#     "caseLevel": None,
#     "entrustType": None,
#     "isHistoryComplaint": None,
#     "repairStatus": None,
#     "isHaveLawsuitOrder": None,
#     "lawsuitType": None,
#     "caseUserUniqueId": None,
#     "caseUserId": None,
#     "departmentIdList": [],
#     "isSensitive": None,
#     "sensitiveTagName": None,
#     "entrustAmountSort": None,
#     "entrustResidueAmountSort": None,
#     "page": 1,
#     "pageSize": 100,
#     "offset": 7928673,
#     "groupByCaseUserUniqueId": 0
# }

# response = requests.post(url, headers=headers, json=payload)

# # 打印结果
# print("Status Code:", response.status_code)
# print("Response JSON:")
# print(response.json())





import requests
import pandas as pd

def fetch_data(token, page):
    url = "https://gateway.fangnuokeji.com/caseCenter/case/allot/orgAllotCaseList"
    headers = {
        "accept": "application/json, text/plain, */*",
        "authorization": 'Bearer ' + token,
        "content-type": "application/json;charset=UTF-8",
        "origin": "https://disposal.fangnuokeji.com",
        "referer": "https://disposal.fangnuokeji.com/",
        "user-agent": "Mozilla/5.0"
    }
    payload = {
        "page": page,
        "pageSize": 100,
        "departmentIdList": [],
        "caseNo": None,
        "productId": None,
        "entrustBatchId": None,
        "batchTypeId": None,
        "userName": None,
        "idno": None,
        "userPhone": None,
        "followStatusId": None,
        "refundStatus": None,
        "caseStatus": None,
        "isRetain": None,
        "retainStagingPlanStatus": None,
        "stagingPlanStatus": None,
        "isTagAlter": None,
        "isFollow": None,
        "entrustContactResultIdList": None,
        "color": None,
        "cpeId": None,
        "startallotTime": None,
        "endallotTime": None,
        "retainEndTimeStart": None,
        "retainEndTimeEnd": None,
        "entrustAmountMax": None,
        "entrustAmountMin": None,
        "regAddrProvince": None,
        "regAddrCity": None,
        "regAddrArea": None,
        "investorName": None,
        "orgTagTempName": None,
        "caseLevelDesc": None,
        "caseLevel": None,
        "entrustType": None,
        "isHistoryComplaint": None,
        "repairStatus": None,
        "isHaveLawsuitOrder": None,
        "lawsuitType": None,
        "caseUserUniqueId": None,
        "caseUserId": None,
        "isSensitive": None,
        "sensitiveTagName": None,
        "entrustAmountSort": None,
        "entrustResidueAmountSort": None,
        "offset": 0,
        "groupByCaseUserUniqueId": 0
    }

    # response = requests.post(url, headers=headers, json=payload)
    response = requests.post(url, headers=headers, json=payload)
    print(f"请求第 {page} 页，状态码: {response.status_code}")
    print("返回内容:", response.text[:500])  # 打印
    if response.status_code != 200:
        print(f"请求第 {page} 页失败，状态码：{response.status_code}")
        return None

    data_json = response.json()
    return data_json.get("data", {}).get("data", [])

def extract_required_fields(item):
    # 临时标签用逗号连接 tagTempList 的 tagName
    temp_tags = ",".join(tag.get("tagName", "") for tag in item.get("tagTempList", []) if tag)
    # 预警标签和风险标签字段示例没给，先空
    warning_tags = ""
    risk_tags = ""

    return {
        "案件ID": item.get("caseNo", ""),
        "产品": item.get("productName", ""),
        "姓名": item.get("userName", ""),
        "证件号": item.get("idno", ""),
        "手机号": item.get("userPhone", ""),
        "委案金额": item.get("entrustAmount", 0),
        "还款入账金额": item.get("handleAmount", 0),
        "减免金额": item.get("entrustReductionAmount", 0),
        "剩余待还金额": item.get("residueAmount", 0),
        "跟进结果": item.get("followStatusText", "") or item.get("entrustContactResultText", ""),
        "处置状态": item.get("caseStatusText", ""),
        "临时标签": temp_tags,
        "预警标签": warning_tags,
        "风险标签": risk_tags,
        "CPE": item.get("cpeName", ""),
        "分案时间": item.get("allotTime", ""),
        "案件状态": item.get("caseStatusText", ""),
        "跟进次数": item.get("entrustFollowTimes", 0),
        "最近跟进时间": item.get("entrustLastFollowTime", ""),
        "债人ID": item.get("caseUserUniqueId", ""),
        "案人ID": item.get("caseUserId", ""),
    }

def main():
    token = input("请输入 Token（Bearer 开头的完整字符串）: ").strip()
    start_page = input("请输入起始页码（例如 1）: ").strip()
    end_page = input("请输入结束页码（例如 3）: ").strip()

    try:
        start_page = int(start_page)
        end_page = int(end_page)
        if start_page > end_page or start_page < 1:
            print("页码范围有误！")
            return
    except:
        print("请输入有效的整数页码！")
        return

    all_data = []
    for page in range(start_page, end_page + 1):
        print(f"正在请求第 {page} 页数据...")
        page_data = fetch_data(token, page)
        if page_data is None:
            print("中断请求。")
            return
        all_data.extend(page_data)

    if not all_data:
        print("未获取到任何数据。")
        return

    processed_data = [extract_required_fields(item) for item in all_data]

    df = pd.DataFrame(processed_data)
    filename = f"exported_data_{start_page}_to_{end_page}.xlsx"
    df.to_excel(filename, index=False)
    print(f"数据导出完成，文件名：{filename}")

if __name__ == "__main__":
    main()
