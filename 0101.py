import requests
import pandas as pd
import time
import random
import urllib3
import datetime
import subprocess
import sys
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# 禁用 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 安全配置区域 =================

# 允许运行的机器码 (如有需要请修改)
ALLOWED_UUIDS = ["ALL"] 

# ================= 常规配置 =================

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

# ================= 基础工具函数 =================

def get_current_machine_uuid():
    try:
        cmd = "wmic csproduct get uuid"
        output = subprocess.check_output(cmd, shell=True).decode()
        lines = output.strip().split('\n')
        return lines[1].strip() if len(lines) >= 2 else "UNKNOWN"
    except: return "ERROR"

def check_permission():
    if "ALL" in ALLOWED_UUIDS: return
    current_uuid = get_current_machine_uuid()
    if current_uuid not in ALLOWED_UUIDS:
        print(f"⛔ 未授权设备 (UUID: {current_uuid})")
        input("按回车键退出...")
        sys.exit()

session = requests.Session()
retries = Retry(total=3, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
session.mount('https://', HTTPAdapter(max_retries=retries))
session.mount('http://', HTTPAdapter(max_retries=retries))

def safe_request(url, params=None):
    try:
        resp = session.get(url, headers=HEADERS, params=params, verify=False, timeout=(5, 10))
        if resp.status_code == 200: return resp.json()
        return None
    except: return None

def clean_case_id(raw_id):
    if not raw_id: return ""
    s_id = str(raw_id)
    if "(" in s_id: return s_id.split("(")[0]
    if "（" in s_id: return s_id.split("（")[0]
    return s_id

# ================= 业务逻辑 =================

def get_detail_data(clean_id):
    time.sleep(random.uniform(0.1, 0.2))
    full_url = f"{DETAIL_BASE_URL}{clean_id}"
    data = safe_request(full_url)
    return data.get("result") if data else {}

def get_plaintext_data(clean_id, type_code):
    time.sleep(random.uniform(0.1, 0.2))
    params = {"id": clean_id, "type": str(type_code)}
    data = safe_request(PLAINTEXT_URL, params)
    return data.get("result", "") if data else ""

def process_record(list_item):
    raw_case_id = list_item.get("caseNo") 
    name = list_item.get("borrowerUserName")
    real_id = clean_case_id(raw_case_id)
    
    print(f" -> 处理: {name} | ID: {real_id} | 请求中...          ", end="\r")
    
    detail = get_detail_data(real_id)
    real_phone = get_plaintext_data(real_id, 1)
    real_id_card = get_plaintext_data(real_id, 2)
    
    def get_val(key, default=""):
        val = detail.get(key)
        if val is not None and str(val) != "": return val
        val = list_item.get(key)
        if val is not None and str(val) != "": return val
        return default

    row_data = {
        "姓名": list_item.get("borrowerUserName"),
        "id": raw_case_id,
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
    print(" 案件数据导出 (自动去重 + 固定排序 + 完整性校验)")
    print("==========================================")
    
    check_permission()

    print("\n")
    input_token = input("请粘贴 Token: ").strip()
    if not input_token: return
    HEADERS["token"] = input_token

    try:
        start_p = int(input("开始页码 (通常填1): "))
        end_p = int(input("结束页码 (例如 120): "))
    except: return

    all_data = []
    # 1. 定义去重集合，用于存放已经抓过的 ID
    seen_ids = set() 
    
    run_id = datetime.datetime.now().strftime("%H%M%S")
    server_total = 0

    for page in range(start_p, end_p + 1):
        print(f"\n====== 正在处理第 {page} 页 ======")
        
        try:
            # 2. 【关键修改】加入 orderByField 固定排序，防止翻页时数据乱跳
            params = {
                "page": str(page), 
                "pageSize": "50", 
                "isAssigned": "1",
                "orderByField": "caseNo", # 强制按案件号排序
                "order": "asc"            # 正序
            }
            
            res = session.get(LIST_URL, headers=HEADERS, params=params, verify=False, timeout=15)
            
            if res.status_code in [401, 403]:
                print("\n❌ Token 已过期！")
                break
            
            data_json = res.json()
            
            # 提取数据和总数
            records = []
            if "result" in data_json:
                records = data_json["result"].get("records", [])
                server_total = data_json["result"].get("total", 0) # 获取服务器总数
            elif "data" in data_json:
                records = data_json["data"].get("records", [])
                server_total = data_json["data"].get("total", 0)

        except Exception as e:
            print(f"列表获取失败: {e}")
            continue

        if not records:
            print("本页无数据。")
            continue

        # 3. 【关键逻辑】循环处理，带去重判断
        page_valid_count = 0
        duplicate_count = 0
        
        for item in records:
            case_id = item.get("caseNo")
            
            # 如果这个 ID 已经在集合里了，说明是重复数据，直接跳过
            if case_id in seen_ids:
                duplicate_count += 1
                continue
            
            # 如果没见过，加入集合，并开始抓取
            seen_ids.add(case_id)
            try:
                row = process_record(item)
                all_data.append(row)
                page_valid_count += 1
            except Exception: pass
        
        # 换行并显示统计
        print(f"\n✅ 第 {page} 页结束: 新增 {page_valid_count} 条，跳过重复 {duplicate_count} 条")
        
        # 临时保存
        try:
            temp_filename = f"临时数据_{start_p}至{page}页_{run_id}.xlsx"
            df = pd.DataFrame(all_data, columns=EXCEL_HEADERS)
            df.to_excel(temp_filename, index=False)
        except: pass

    # 4. 最终结果核对
    print("\n" + "="*40)
    print(f"统计报告:")
    print(f"服务器显示总数: {server_total} 条")
    print(f"实际抓取去重后: {len(all_data)} 条")
    if server_total > 0:
        completion_rate = (len(all_data) / server_total) * 100
        print(f"数据完整率: {completion_rate:.2f}%")
    print("="*40)

    if all_data:
        final_filename = f"最终导出_{start_p}-{end_p}页_{run_id}.xlsx"
        try:
            pd.DataFrame(all_data, columns=EXCEL_HEADERS).to_excel(final_filename, index=False)
            print(f"\n🎉 成功保存: {final_filename}")
        except: pass
        
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()
