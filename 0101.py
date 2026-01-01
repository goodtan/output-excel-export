import requests
import pandas as pd
import time
import random
import urllib3
import datetime

# 禁用 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 配置区域 =================

# 1. 列表接口
LIST_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/query"
# 2. 详情接口 (拼接 ID)
DETAIL_BASE_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/find/"
# 3. 明文敏感信息接口 (需要参数 id 和 type)
PLAINTEXT_URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/show/plaintext"

# 基础请求头
HEADERS = {
    "accept": "application/json, text/plain, */*",
    "referer": "https://kuafu.dadixintong.com/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
    # "token":  <--- 运行时动态输入
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
    获取详情页数据
    """
    full_url = f"{DETAIL_BASE_URL}{case_id}"
    try:
        # 极短延时
        time.sleep(random.uniform(0.1, 0.3)) 
        resp = requests.get(full_url, headers=HEADERS, verify=False, timeout=8)
        if resp.status_code == 200:
            res_json = resp.json()
            return res_json.get("result") or {}
        return {}
    except Exception:
        return {}

def get_plaintext_data(case_id, type_code):
    """
    获取明文数据 (身份证或手机号)
    type_code: 1=手机号, 2=身份证
    """
    params = {
        "id": case_id,
        "type": str(type_code)
    }
    try:
        # 每次请求稍微停顿一下，防止并发过高
        time.sleep(random.uniform(0.1, 0.3))
        
        resp = requests.get(PLAINTEXT_URL, headers=HEADERS, params=params, verify=False, timeout=8)
        if resp.status_code == 200:
            res_json = resp.json()
            # 返回 result 字段里的字符串
            return res_json.get("result", "")
        return ""
    except Exception as e:
        print(f"   [明文获取失败 type={type_code}] ID:{case_id} {e}")
        return ""

def process_record(list_item):
    """
    核心处理函数：列表 + 详情 + 明文手机 + 明文身份证
    """
    case_id = list_item.get("caseNo") 
    name = list_item.get("borrowerUserName")
    
    print(f" -> 处理: {name} | 正在获取详情及敏感信息...", end="\r")
    
    # 1. 获取详情页数据
    detail = get_detail_data(case_id)
    
    # 2. 获取明文手机号 (type=1)
    real_phone = get_plaintext_data(case_id, 1)
    
    # 3. 获取明文身份证 (type=2)
    real_id_card = get_plaintext_data(case_id, 2)
    
    # 4. 辅助取值函数：优先从 detail 取，没有则从 list_item 取
    #    (根据你的要求，列表有的从列表取，列表没有找详情)
    def get_val(key, default=""):
        # 优先看详情里有没有
        val = detail.get(key)
        if val is not None and str(val) != "":
            return val
        # 详情没有，看列表里有没有
        val = list_item.get(key)
        if val is not None and str(val) != "":
            return val
        return default

    # 5. 组装数据
    row_data = {
        "姓名": list_item.get("borrowerUserName"), # 优先用列表的
        "id": case_id,
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
        
        # --- 重点：使用明文接口的数据 ---
        "证件号": real_id_card if real_id_card else get_val("borrowerIdCard"),
        "本人手机号码": real_phone if real_phone else get_val("borrowerTel"),
        # -----------------------------
        
        "所在部门": get_val("deptName"), # 详情里的字段
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
        "户籍地址": get_val("residenceAddress"), # 详情里的字段
        "电话信息": get_val("telLatestTime"),    # 列表里的字段
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
    print(f" -> 处理: {name} | ✅ 数据获取完毕                 ")
    return row_data

def main():
    print("==========================================")
    print("   案件全量数据导出 (含明文手机/身份证)")
    print("==========================================\n")

    input_token = input("请粘贴最新的 Token 并按回车: ").strip()
    if not input_token:
        print("错误：Token 不能为空")
        input("按回车键退出...")
        return
    HEADERS["token"] = input_token
    print("✅ Token 已设置！\n")

    try:
        start_p = int(input("请输入开始页码: "))
        end_p = int(input("请输入结束页码: "))
    except ValueError:
        print("输入错误")
        return

    all_data = []
    # 运行 ID，防止文件名冲突
    run_id = datetime.datetime.now().strftime("%H%M%S")

    for page in range(start_p, end_p + 1):
        print(f"\n====== 正在处理第 {page} 页 ======")
        
        try:
            params = {"page": str(page), "pageSize": "50", "isAssigned": "1"}
            # 列表请求
            res = requests.get(LIST_URL, headers=HEADERS, params=params, verify=False, timeout=15)
            
            if res.status_code in [401, 403]:
                print("\n❌ Token 已过期，请重新获取！")
                break
            
            if res.status_code != 200:
                print(f"列表请求失败: {res.status_code}")
                continue
                
            data_json = res.json()
            if "result" in data_json and "records" in data_json["result"]:
                records = data_json["result"]["records"]
            elif "data" in data_json and "records" in data_json["data"]:
                records = data_json["data"]["records"]
            else:
                records = []
                
        except Exception as e:
            print(f"网络请求错误: {e}")
            continue

        if not records:
            print("本页无数据。")
            continue

        # 循环处理每一条
        for item in records:
            try:
                row = process_record(item)
                all_data.append(row)
            except Exception as e:
                print(f"\n⚠️ 跳过异常数据: {e}")
                continue
        
        # 临时保存
        print(f"第 {page} 页完成，正在保存...")
        try:
            temp_filename = f"临时数据_{start_p}至{page}页_{run_id}.xlsx"
            df = pd.DataFrame(all_data, columns=EXCEL_HEADERS)
            df.to_excel(temp_filename, index=False)
            print(f"✅ 已保存: {temp_filename}")
        except Exception as e:
            print(f"保存失败: {e}")

    # 最终保存
    print("\n------------------------------------------")
    if all_data:
        final_filename = f"案件导出_{start_p}-{end_p}页_全量_{run_id}.xlsx"
        try:
            df = pd.DataFrame(all_data, columns=EXCEL_HEADERS)
            df.to_excel(final_filename, index=False)
            print(f"🎉 成功！文件已生成: {final_filename}")
        except Exception as e:
            print(f"❌ 最终保存失败: {e}")
            df.to_excel(f"backup_{run_id}.xlsx", index=False)
    else:
        print("未获取到数据")
        
    input("\n程序运行结束，按回车键退出...")

if __name__ == "__main__":
    main()
