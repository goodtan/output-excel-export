import requests
import pandas as pd
import time
import random
import urllib3
import datetime
import os

# 禁用 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 配置区域 =================

# 接口地址
LIST_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/query"
# 详情接口 (注意：ID将直接拼接到此URL后面)
DETAIL_BASE_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/find/"

# 基础请求头
HEADERS = {
    "accept": "application/json, text/plain, */*",
    "referer": "https://kuafu.dadixintong.com/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
    # "token":  <--- 待会儿在 main 函数里动态添加
}

# Excel 表头
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

# ================= 核心逻辑 =================

def get_detail_data(case_id):
    """
    根据 ID 获取详情
    """
    # 直接拼接 URL，适配带 (E) 的 ID
    full_url = f"{DETAIL_BASE_URL}{case_id}"
    
    try:
        # 随机休眠，防止封号
        time.sleep(random.uniform(0.2, 0.5))
        
        # 发送请求
        resp = requests.get(full_url, headers=HEADERS, verify=False, timeout=10)
        
        if resp.status_code == 200:
            res_json = resp.json()
            # 【关键修复】如果 result 是 None，返回空字典 {}
            return res_json.get("result") or {}
        else:
            print(f"   [详情失败] ID:{case_id} 状态码:{resp.status_code}")
            return {}
            
    except Exception as e:
        print(f"   [详情异常] ID:{case_id} 错误:{e}")
        return {}

def process_record(list_item):
    """
    处理单条数据：合并列表和详情
    """
    case_id = list_item.get("caseNo") 
    name = list_item.get("borrowerUserName")
    
    print(f" -> 正在抓取: {name} ({case_id})")
    
    # 1. 获取详情
    detail = get_detail_data(case_id)
    
    # 2. 【关键修复】安全获取字段函数
    def get_val(data_dict, key):
        # 如果源数据本身是 None，返回空字符串
        if data_dict is None:
            return ""
        val = data_dict.get(key)
        # 如果获取到的值是 None，返回空字符串
        return val if val is not None else ""

    # 3. 字段映射 (数据组装)
    row_data = {
        "姓名": get_val(list_item, "borrowerUserName"),
        "id": case_id,
        "案件类型": get_val(detail, "caseStage"), 
        "借款金额": get_val(detail, "financeAmount"),
        "逾期期数": f"{get_val(detail, 'financeOverdueStart')}-{get_val(detail, 'financeOverdueEnd')}",
        "跟进人": get_val(detail, "followName"),
        "产品名称": get_val(detail, "productName"),
        "渠道APP名称": get_val(detail, "showCompanyInfo"), # 对应 "易得花"
        "全部结清": "", 
        "待还最大逾期天数": get_val(detail, "financeOverdueDays"),
        "提前结清": "",
        "剩余应还本金": get_val(detail, "leftNeedRepayPrincipal"),
        "剩余应还利息": get_val(detail, "leftNeedRepayInterest"),
        "所在省市": get_val(detail, "borrowerArea"),
        "证件号": get_val(detail, "borrowerIdCard"),
        "本人手机号码": get_val(detail, "borrowerTel"),
        "所在部门": get_val(detail, "deptName"), 
        "贷后逾期天数": get_val(detail, "reminderOverdueDays"),
        "资金方代码": get_val(detail, "fundSideCode"),
        "进件渠道": get_val(detail, "productChannel"),
        "逾期加当期": get_val(detail, "settleAmount"),
        "期限": get_val(detail, "totalPeriod"),
        "借款日期": get_val(detail, "financeLoanTime"),
        "只还全部逾期": get_val(detail, "totalOverdueAmount"),
        "代收逾期费": get_val(detail, "needRepayOverdueFeeAmount"),
        "借款标的": get_val(detail, "bidId"),
        "借款年利率": get_val(detail, "apr"),
        "户籍地址": get_val(detail, "residenceAddress"),
        "电话信息": get_val(list_item, "telLatestTime"),
        "客诉类型": "",
        "客诉内容": "",
        "协商方案": "",
        "跟进记录": "",
        "反馈时间": "",
        "处理人": "",
        "对应工单编号": "",
        "应还金额": get_val(detail, "financeNeedRepayTotal"),
        "实收金额": get_val(detail, "receivedAmount"),
        "代收金额": ""
    }
    return row_data

def main():
    print("==========================================")
    print("     案件数据导出工具 (终极稳定版)")
    print("==========================================\n")

    # 1. 获取 Token
    input_token = input("请粘贴最新的 Token 并按回车: ").strip()
    if not input_token:
        print("错误：Token 不能为空！")
        input("按回车键退出...")
        return
    HEADERS["token"] = input_token
    print("✅ Token 已设置！\n")

    # 2. 获取页码
    try:
        start_p = int(input("请输入开始页码 (例如 1): "))
        end_p = int(input("请输入结束页码 (例如 5): "))
    except ValueError:
        print("输入错误，请输入纯数字")
        input("按回车键退出...")
        return

    all_data = []
    
    # 3. 生成运行ID (时间戳)，防止文件名冲突
    # 例如：run_143005 (14点30分05秒)
    run_id = datetime.datetime.now().strftime("%H%M%S")
    print(f"本次运行 ID: {run_id} (用于生成唯一文件名)\n")

    # 4. 开始循环
    for page in range(start_p, end_p + 1):
        print(f"\n====== 正在处理第 {page} 页 ======")
        
        # --- 获取列表 ---
        try:
            params = {"page": str(page), "pageSize": "50", "isAssigned": "1"}
            res = requests.get(LIST_URL, headers=HEADERS, params=params, verify=False, timeout=15)
            
            # 检查 Token 是否过期
            if res.status_code in [401, 403]:
                print("\n❌ 严重错误：Token 已过期！请重新去浏览器复制最新的 Token。")
                break
            
            if res.status_code != 200:
                print(f"列表请求失败，状态码: {res.status_code}")
                continue
                
            data_json = res.json()
            
            # 兼容 result 或 data 字段
            if "result" in data_json and "records" in data_json["result"]:
                records = data_json["result"]["records"]
            elif "data" in data_json and "records" in data_json["data"]:
                records = data_json["data"]["records"]
            else:
                records = []
                
        except Exception as e:
            print(f"列表请求发生网络错误: {e}")
            continue

        if not records:
            print(f"第 {page} 页没有数据，跳过。")
            continue

        # --- 获取详情 ---
        for item in records:
            try:
                row = process_record(item)
                all_data.append(row)
            except Exception as e:
                # 即使某一条出错，也不要崩溃，打印错误并继续
                print(f"⚠️ 跳过异常数据: {e}")
                continue
        
        # --- 临时保存 (每页存一次) ---
        print(f"第 {page} 页完成，正在临时保存...")
        try:
            # 文件名带上 run_id，避免 Permission denied
            temp_filename = f"临时数据_{start_p}至{page}页_{run_id}.xlsx"
            df = pd.DataFrame(all_data, columns=EXCEL_HEADERS)
            df.to_excel(temp_filename, index=False)
            print(f"✅ 已保存: {temp_filename}")
        except Exception as e:
            print(f"⚠️ 临时保存失败 (不影响继续运行): {e}")

    # 5. 最终保存
    print("\n------------------------------------------")
    if all_data:
        final_filename = f"案件导出_{start_p}-{end_p}页_{run_id}.xlsx"
        try:
            df = pd.DataFrame(all_data, columns=EXCEL_HEADERS)
            df.to_excel(final_filename, index=False)
            print(f"🎉 成功！最终文件已生成: {final_filename}")
        except Exception as e:
            print(f"❌ 最终保存失败: {e}")
            # 备用保存方案
            fallback_name = f"data_backup_{run_id}.xlsx"
            df.to_excel(fallback_name, index=False)
            print(f"已尝试保存为备用文件: {fallback_name}")
    else:
        print("⚠️ 本次运行未获取到任何数据。")
        
    input("\n程序运行结束，请按回车键退出...")

if __name__ == "__main__":
    main()
