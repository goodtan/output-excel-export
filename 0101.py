import requests
import pandas as pd
import time
import random
import urllib3

# 禁用 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 配置区域 =================

# 接口地址
LIST_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/query"
DETAIL_BASE_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/find/"

# 基础请求头
HEADERS = {
    "accept": "application/json, text/plain, */*",
    "referer": "https://kuafu.dadixintong.com/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
    # "token":  等待用户输入
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
    # 修复：如果 ID 带有 (E) 等后缀，可能导致接口查不到，这里尝试原样请求
    full_url = f"{DETAIL_BASE_URL}{case_id}"
    
    try:
        time.sleep(random.uniform(0.3, 0.6))
        # 增加 verify=False 和超时设置
        resp = requests.get(full_url, headers=HEADERS, verify=False, timeout=10)
        
        if resp.status_code == 200:
            res_json = resp.json()
            # 【关键修复】如果 result 是 None，返回空字典 {}，防止后续报错
            return res_json.get("result") or {}
        else:
            print(f"   [详情失败] ID:{case_id} 状态码:{resp.status_code}")
            return {} # 返回空字典
            
    except Exception as e:
        print(f"   [详情异常] ID:{case_id} 错误:{e}")
        return {} # 返回空字典

def process_record(list_item):
    """
    合并列表数据和详情数据
    """
    case_id = list_item.get("caseNo") 
    name = list_item.get("borrowerUserName")
    
    print(f" -> 正在抓取详情: {name} (ID: {case_id})")
    
    # 1. 请求详情
    detail = get_detail_data(case_id)
    
    # 2. 辅助函数 【关键修复】
    def get_val(data_dict, key):
        # 如果传入的数据本身是 None，直接返回空字符串
        if data_dict is None:
            return ""
        val = data_dict.get(key)
        return val if val is not None else ""

    # 3. 字段映射
    row_data = {
        "姓名": get_val(list_item, "borrowerUserName"),
        "id": case_id,
        "案件类型": get_val(detail, "caseStage"), 
        "借款金额": get_val(detail, "financeAmount"),
        "逾期期数": f"{get_val(detail, 'financeOverdueStart')}-{get_val(detail, 'financeOverdueEnd')}",
        "跟进人": get_val(detail, "followName"),
        "产品名称": get_val(detail, "productName"),
        "渠道APP名称": get_val(detail, "showCompanyInfo"), 
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
    print("     案件数据导出工具 (防崩溃版)")
    print("==========================================\n")

    input_token = input("请粘贴最新的 Token 并按回车: ").strip()
    
    if not input_token:
        print("错误：Token 不能为空！")
        input("按回车键退出...")
        return

    HEADERS["token"] = input_token
    print("✅ Token 已设置！\n")

    try:
        start_p = int(input("请输入开始页码: "))
        end_p = int(input("请输入结束页码: "))
    except:
        print("输入错误，请输入数字")
        input("按回车键退出...")
        return

    all_data = []

    for page in range(start_p, end_p + 1):
        print(f"\n====== 正在处理第 {page} 页 ======")
        
        try:
            params = {"page": str(page), "pageSize": "50", "isAssigned": "1"}
            res = requests.get(LIST_URL, headers=HEADERS, params=params, verify=False, timeout=10)
            
            if res.status_code in [401, 403]:
                print("❌ 错误：Token 已过期或无效，请重新抓取 Token。")
                break
                
            if res.status_code != 200:
                print(f"列表页请求失败，状态码: {res.status_code}")
                continue
                
            data_json = res.json()
            
            if "result" in data_json and "records" in data_json["result"]:
                records = data_json["result"]["records"]
            elif "data" in data_json and "records" in data_json["data"]:
                records = data_json["data"]["records"]
            else:
                print("列表数据结构异常或无数据")
                records = []
                
        except Exception as e:
            print(f"列表请求出错: {e}")
            continue

        if not records:
            print("本页无数据。")
            continue

        for item in records:
            # 增加 try-except 保护，防止某一条数据异常导致整个程序闪退
            try:
                row = process_record(item)
                all_data.append(row)
            except Exception as e:
                print(f"⚠️ 跳过异常数据 {item.get('borrowerUserName', '未知')}: {e}")
                continue
            
        print(f"第 {page} 页数据已获取，正在临时保存...")
        try:
            temp_df = pd.DataFrame(all_data, columns=EXCEL_HEADERS)
            temp_df.to_excel(f"temp_data_page_{start_p}_to_{page}.xlsx", index=False)
        except Exception as e:
            print(f"临时保存失败: {e}")

    print("\n✅ 所有任务完成！")
    
    if all_data:
        final_filename = f"案件明细导出_{start_p}-{end_p}页.xlsx"
        try:
            df = pd.DataFrame(all_data, columns=EXCEL_HEADERS)
            df.to_excel(final_filename, index=False)
            print(f"🎉 最终文件已生成: {final_filename}")
        except Exception as e:
            print(f"❌ 保存文件失败 (请检查文件是否被占用): {e}")
    else:
        print("⚠️ 未获取到任何数据")
        
    input("\n程序运行结束，按回车键退出...")

if __name__ == "__main__":
    main()
