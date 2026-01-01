import requests
import pandas as pd
import time
import random
import urllib3
import datetime
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# 禁用 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 配置区域 =================

LIST_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/query"
DETAIL_BASE_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/find/"
PLAINTEXT_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/show/plaintext"

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "referer": "https://kuafu.dadixintong.com/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
    # "token":  <--- 运行时输入
}

EXCEL_HEADERS = [
    "姓名", "id", "案件类型", "借款金额", "逾期期数", "跟进人", 
    "产品名称", "渠道APP名称", "全部结清", "待还最大逾期天数", "提前结清", 
    "剩余应还本金", "剩余应还利息", "所在省市", "证件号", "本人手机号码", 
    "所在部门", "贷后逾期天数", "资金方代码", "进件渠道", "逾期加当期", 
    "期限", "借款日期", "只还全部逾期", "代收逾期费", "借款标的", 
    "借款年利率", "户籍地址", "电话信息", "客诉类型", "客诉内容", 
    "协商方案", "跟进记录", "反馈时间", "处理人", "对应工单编号", 
    "应还金额", "实收金额", "代收金额"
]

# ================= 网络请求增强版 (Session) =================

session = requests.Session()
retries = Retry(total=3, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
session.mount('https://', HTTPAdapter(max_retries=retries))
session.mount('http://', HTTPAdapter(max_retries=retries))

def safe_request(url, params=None):
    """
    安全请求函数：自动重试，防卡死
    """
    try:
        # 5秒连接超时，10秒读取超时
        resp = session.get(url, headers=HEADERS, params=params, verify=False, timeout=(5, 10))
        if resp.status_code == 200:
            return resp.json()
        return None
    except Exception:
        return None

# ================= 辅助工具 =================

def clean_case_id(raw_id):
    """
    【关键修正】清洗 ID
    将 "1420963568373014574(E)" 变成 "1420963568373014574"
    """
    if not raw_id:
        return ""
    # 转成字符串
    s_id = str(raw_id)
    # 如果包含左括号，只取左括号前面的部分
    if "(" in s_id:
        return s_id.split("(")[0]
    if "（" in s_id: # 兼容中文括号
        return s_id.split("（")[0]
    return s_id

# ================= 核心逻辑 =================

def get_detail_data(clean_id):
    """请求详情 (使用清洗后的 ID)"""
    time.sleep(random.uniform(0.1, 0.2))
    full_url = f"{DETAIL_BASE_URL}{clean_id}"
    data = safe_request(full_url)
    return data.get("result") if data else {}

def get_plaintext_data(clean_id, type_code):
    """请求明文数据 (使用清洗后的 ID)"""
    time.sleep(random.uniform(0.1, 0.2))
    params = {"id": clean_id, "type": str(type_code)}
    data = safe_request(PLAINTEXT_URL, params)
    return data.get("result", "") if data else ""

def process_record(list_item):
    # 1. 获取原始 ID 和 名字
    raw_case_id = list_item.get("caseNo") 
    name = list_item.get("borrowerUserName")
    
    # 2. 【关键】获取清洗后的 ID (去掉 (E))
    real_id = clean_case_id(raw_case_id)
    
    # 打印进度 (用 \r 实现不换行刷新)
    print(f" -> 处理: {name} | ID: {real_id} | 正在请求...          ", end="\r")
    
    # 3. 使用清洗后的 ID 去请求各个接口
    detail = get_detail_data(real_id)
    real_phone = get_plaintext_data(real_id, 1)   # 获取手机
    real_id_card = get_plaintext_data(real_id, 2) # 获取身份证
    
    # 4. 辅助取值
    def get_val(key, default=""):
        val = detail.get(key)
        if val is not None and str(val) != "": return val
        val = list_item.get(key)
        if val is not None and str(val) != "": return val
        return default

    # 5. 组装数据
    row_data = {
        "姓名": list_item.get("borrowerUserName"),
        "id": raw_case_id, # Excel 里保留原始带(E)的ID，方便核对
        "案件类型": get_val("caseStage"), 
        "借款金额": get_val("financeAmount"),
        "逾期期数": f"{get_val('financeOverdueStart')}-{get_val('financeOverdueEnd')}",
        "跟进人": get_val("followName"),
        "产品名称": get_val("productName"),
        "渠道APP名称": get_val("showCompanyInfo"), 
        "全部结清": "", 
        "待还最大逾期天数": get_val("financeOverdueDays"),
        "提前结清": "",
        "剩余应还本金": get_val("leftNeedRepayPrincipal"),
        "剩余应还利息": get_val("leftNeedRepayInterest"),
        "所在省市": get_val("borrowerArea"),
        # 使用明文接口数据
        "证件号": real_id_card if real_id_card else get_val("borrowerIdCard"),
        "本人手机号码": real_phone if real_phone else get_val("borrowerTel"),
        "所在部门": get_val("deptName"), 
        "贷后逾期天数": get_val("reminderOverdueDays"),
        "资金方代码": get_val("fundSideCode"),
        "进件渠道": get_val("productChannel"),
        "逾期加当期": get_val("settleAmount"),
        "期限": get_val("totalPeriod"),
        "借款日期": get_val("financeLoanTime"),
        "只还全部逾期": get_val("totalOverdueAmount"),
        "代收逾期费": get_val("needRepayOverdueFeeAmount"),
        "借款标的": get_val("bidId"),
        "借款年利率": get_val("apr"),
        "户籍地址": get_val("residenceAddress"),
        "电话信息": get_val("telLatestTime"),
        "客诉类型": "",
        "客诉内容": "",
        "协商方案": "",
        "跟进记录": "",
        "反馈时间": "",
        "处理人": "",
        "对应工单编号": "",
        "应还金额": get_val("financeNeedRepayTotal"),
        "实收金额": get_val("receivedAmount"),
        "代收金额": ""
    }
    return row_data

def main():
    print("==========================================")
    print(" 案件数据导出 (自动清洗ID后缀 + 防卡死)")
    print("==========================================\n")

    input_token = input("请粘贴 Token: ").strip()
    if not input_token: return
    HEADERS["token"] = input_token

    try:
        start_p = int(input("开始页码: "))
        end_p = int(input("结束页码: "))
    except: return

    all_data = []
    run_id = datetime.datetime.now().strftime("%H%M%S")

    for page in range(start_p, end_p + 1):
        print(f"\n====== 正在处理第 {page} 页 ======")
        
        try:
            params = {"page": str(page), "pageSize": "50", "isAssigned": "1"}
            res = session.get(LIST_URL, headers=HEADERS, params=params, verify=False, timeout=15)
            
            if res.status_code in [401, 403]:
                print("\n❌ Token 已过期！")
                break
            
            data_json = res.json()
            if "result" in data_json and "records" in data_json["result"]:
                records = data_json["result"]["records"]
            elif "data" in data_json and "records" in data_json["data"]:
                records = data_json["data"]["records"]
            else:
                records = []
        except Exception as e:
            print(f"列表获取失败: {e}")
            continue

        if not records:
            print("本页无数据。")
            continue

        for item in records:
            try:
                row = process_record(item)
                all_data.append(row)
            except Exception: pass
        
        # 换行
        print(f"\n第 {page} 页完成，保存中...")
        
        try:
            temp_filename = f"临时数据_{start_p}至{page}页_{run_id}.xlsx"
            df = pd.DataFrame(all_data, columns=EXCEL_HEADERS)
            df.to_excel(temp_filename, index=False)
            print(f"✅ 已保存: {temp_filename}")
        except: pass

    if all_data:
        final_filename = f"最终导出_{run_id}.xlsx"
        try:
            pd.DataFrame(all_data, columns=EXCEL_HEADERS).to_excel(final_filename, index=False)
            print(f"\n🎉 完成！文件: {final_filename}")
        except: pass
        
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()
