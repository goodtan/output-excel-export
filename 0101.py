import requests
import pandas as pd
import time
import json
import urllib3

# 禁用 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 配置区域 =================
# 请务必在此处填入最新的 Token，否则会报 403
TOKEN = "2005569195271577602:01kdw1fp6j1fb06mzzggy6b66y"  # <--- 在这里更新 Token

# 接口地址
URL = "https://kuafu.dadixintong.com/reminder-app/cases/case/query"

# Excel 表头定义 (根据你的要求)
EXCEL_HEADERS = [
    "姓名", "id", "案件类型", "借款金额", "逾期期数", "跟进人", "产品名称", "渠道APP名称",
    "全部结清", "待还最大逾期天数", "提前结清", "剩余应还本金", "剩余应还利息", "所在省市",
    "证件号", "本人手机号码", "所在部门", "贷后逾期天数", "资金方代码", "进件渠道",
    "逾期加当期", "期限", "借款日期", "只还全部逾期", "代收逾期费", "借款标的",
    "借款年利率", "户籍地址", "电话信息", "客诉类型", "客诉内容", "协商方案",
    "跟进记录", "反馈时间", "处理人", "对应工单编号", "应还金额", "实收金额", "代收金额"
]

# 请求头 (已优化，防止 403)
HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "zh-CN,zh;q=0.9",
    "referer": "https://kuafu.dadixintong.com/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
    "token": TOKEN
}

def fetch_page_data(page_num, page_size=50):
    """
    请求单页数据
    """
    params = {
        "page": str(page_num),
        "pageSize": str(page_size),
        "isAssigned": "1" # 根据你之前的 cURL 保留此参数
        # 其他筛选参数如果需要可以在这里添加
    }
    
    try:
        response = requests.get(URL, headers=HEADERS, params=params, verify=False, timeout=10)
        if response.status_code == 200:
            res_json = response.json()
            # 兼容处理：检查数据是在 result.records 还是 data.records
            if "result" in res_json and res_json["result"] and "records" in res_json["result"]:
                return res_json["result"]["records"]
            elif "data" in res_json and "records" in res_json["data"]:
                return res_json["data"]["records"]
            else:
                print(f"警告：第 {page_num} 页返回结构异常: {res_json.keys()}")
                return []
        else:
            print(f"请求失败 第 {page_num} 页: 状态码 {response.status_code}")
            print("错误详情:", response.text)
            return None
    except Exception as e:
        print(f"请求异常 第 {page_num} 页: {e}")
        return None

def process_record(item):
    """
    将 API 返回的单条 JSON 数据映射到 Excel 表头
    """
    # 辅助函数：安全获取字段，如果为 None 返回空字符串
    def get_val(key, default=""):
        return item.get(key) if item.get(key) is not None else default

    return {
        "姓名": get_val("borrowerUserName"),
        "id": get_val("caseNo"), # 既然 caseNo 和 id 一样，优先用 caseNo
        "案件类型": get_val("caseStage"), # 或者是 loanType，根据业务调整
        "借款金额": get_val("financeAmount"),
        "逾期期数": f"{get_val('financeOverdueStart')}-{get_val('financeOverdueEnd')}",
        "跟进人": get_val("followName"),
        "产品名称": get_val("productName"),
        "渠道APP名称": "", # JSON 中未找到
        "全部结清": "", # JSON 中未找到 (可能是 calculated field)
        "待还最大逾期天数": get_val("financeOverdueDays"),
        "提前结清": "", # JSON 中未找到
        "剩余应还本金": get_val("leftNeedRepayPrincipal"),
        "剩余应还利息": "", # JSON 中只有 totalOverdueAmount (总逾期) 和 principal
        "所在省市": "", # JSON 中未找到
        "证件号": get_val("borrowerIdCard"),
        "本人手机号码": get_val("borrowerTel"),
        "所在部门": get_val("deptId"), # 这里是 ID，没有部门名称
        "贷后逾期天数": get_val("reminderOverdueDays"),
        "资金方代码": get_val("fundSideCode"),
        "进件渠道": "", # JSON 中未找到
        "逾期加当期": "", 
        "期限": get_val("totalPeriod"),
        "借款日期": get_val("financeLoanTime"),
        "只还全部逾期": "",
        "代收逾期费": "",
        "借款标的": get_val("bidId"),
        "借款年利率": "", # JSON 中未找到
        "户籍地址": "", # JSON 中未找到
        "电话信息": get_val("telLatestTime"), # 映射为最近通话时间
        "客诉类型": "",
        "客诉内容": "",
        "协商方案": "",
        "跟进记录": "",
        "反馈时间": "", # 可能是 editTime?
        "处理人": "",
        "对应工单编号": "",
        "应还金额": get_val("financeNeedRepayTotal"),
        "实收金额": "",
        "代收金额": ""
    }

def main():
    print("=== 数据导出工具 ===")
    
    try:
        start_page = int(input("请输入开始页码 (例如 1): "))
        end_page = int(input("请输入结束页码 (例如 5): "))
    except ValueError:
        print("页码必须是数字！")
        return

    all_rows = []
    
    print(f"\n开始抓取，从第 {start_page} 页 到 第 {end_page} 页...\n")

    for page in range(start_page, end_page + 1):
        print(f"正在请求第 {page} 页...")
        records = fetch_page_data(page)
        
        if records is None:
            print("遇到错误，停止抓取。")
            break
            
        if not records:
            print(f"第 {page} 页无数据，可能已经到底了。")
            break
            
        # 处理当前页数据
        for item in records:
            row_data = process_record(item)
            all_rows.append(row_data)
            
        # 暂停一下，防止请求过快被封 IP
        time.sleep(1)

    # 保存到 Excel
    if all_rows:
        print(f"\n共获取到 {len(all_rows)} 条数据，正在写入 Excel...")
        df = pd.DataFrame(all_rows, columns=EXCEL_HEADERS)
        
        filename = f"案件数据_{start_page}页至{end_page}页.xlsx"
        df.to_excel(filename, index=False)
        print(f"✅ 成功！文件已保存为: {filename}")
    else:
        print("⚠️ 未获取到任何数据。")

if __name__ == "__main__":
    main()
