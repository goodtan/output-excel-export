import time
import pandas as pd
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoAlertPresentException, UnexpectedAlertPresentException, TimeoutException

# --- 配置 ---
EXCEL_FILE = '借贷人数据.xlsx'

# 基础登录地址 (只填 IP 和端口即可，系统会自动跳登录页)
BASE_URL = "http://10.200.18.179:8088" 

def get_browser():
    """启动 Google Chrome 浏览器"""
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-certificate-errors')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
        
    driver_path = os.path.join(base_path, "chromedriver.exe")
    
    if not os.path.exists(driver_path):
        print("❌ 错误：找不到 chromedriver.exe")
        input("按回车键退出...")
        sys.exit(1)

    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def save_to_excel(data_dict):
    df_new = pd.DataFrame([data_dict])
    if not os.path.exists(EXCEL_FILE):
        df_new.to_excel(EXCEL_FILE, index=False)
    else:
        df_old = pd.read_excel(EXCEL_FILE)
        df_all = pd.concat([df_old, df_new], ignore_index=True)
        df_all.to_excel(EXCEL_FILE, index=False)
    print(f"✅ 已保存: {data_dict.get('姓名')}")

def smart_switch_to_iframe(driver):
    """智能寻找包含数据的 iframe"""
    print("🔍 正在扫描当前页面寻找数据...")
    
    # 1. 先切回最外层
    driver.switch_to.default_content()
    
    # 2. 优先尝试根据 ID 切换 (根据之前的源码分析，这是最准的)
    try:
        driver.switch_to.frame("frmcaseMainInfo")
        # 验证一下里面有没有 Borrow借贷人
        if len(driver.find_elements(By.XPATH, "//*[contains(text(), 'Borrower借贷人')]")) > 0:
            print("✅ 已切入数据框架 (frmcaseMainInfo)")
            return True
        driver.switch_to.default_content() # 不对就退出来
    except:
        pass

    # 3. 如果 ID 不对，暴力遍历所有 iframe
    iframe_list = driver.find_elements(By.TAG_NAME, "iframe")
    for i in range(len(iframe_list)):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(i)
            if len(driver.find_elements(By.XPATH, "//*[contains(text(), 'Borrower借贷人')]")) > 0:
                print(f"✅ 在第 {i+1} 个框架中找到数据")
                return True
        except:
            continue
            
    print("❌ 当前页面未找到数据！(请确认是否在【详情页】)")
    return False

def main():
    try:
        driver = get_browser()
    except Exception as e:
        print(f"❌ 浏览器启动失败: {e}")
        input("按回车键退出...")
        return

    # 1. 打开基础首页
    driver.get(BASE_URL)

    # 2. 【关键】等待用户手动操作
    print("\n" + "="*60)
    print("【请手动操作浏览器】")
    print("1. 输入账号密码登录系统。")
    print("2. 点击菜单，进入任务列表。")
    print("3. 点击【第一个客户】，进入它的【详情页】（能看到电话号码的页面）。")
    print("4. 确保页面加载完毕后，回到这里，按【回车键】开始自动抓取。")
    print("="*60 + "\n")
    input(">> 准备好后，按回车键开始: ") 

    count = 0
    is_finished = False

    while not is_finished:
        try:
            print(f"\n>> 正在处理第 {count + 1} 个任务...")
            
            # --- 步骤1：智能寻找并切入 iframe ---
            if not smart_switch_to_iframe(driver):
                print("⚠️ 无法定位数据，重试中...")
                time.sleep(3)
                # 如果还是找不到，可能是页面还在加载，或者已经跳出去了
                # 这里选择 continue 重试，或者你可以选择抛出异常
                continue

            wait = WebDriverWait(driver, 10)
            
            # --- 步骤2：定位行 ---
            borrower_tr = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//td[contains(text(), 'Borrower')]/..")
            ))

            # --- 步骤3：点击显示 ---
            try:
                show_btn = borrower_tr.find_element(By.XPATH, ".//a[contains(text(), '显示')]")
                driver.execute_script("arguments[0].click();", show_btn)
                time.sleep(0.8) 
            except:
                pass 

            # --- 步骤4：提取数据 ---
            cols = borrower_tr.find_elements(By.TAG_NAME, "td")
            phone_cell = cols[2]
            try:
                phone_text = phone_cell.find_element(By.XPATH, ".//a[contains(@id, 'phoneValue')]").text
            except:
                phone_text = phone_cell.text.replace("显示", "").strip()

            data = {
                "姓名": cols[0].text,
                "关系": cols[1].text,
                "电话号码": phone_text,
                "号码来源": cols[3].text,
                "电话类型": cols[4].text,
                "是否有效": cols[5].text,
                "备注": cols[6].text
            }
            save_to_excel(data)

            # --- 步骤5：切回主页面操作按钮 ---
            driver.switch_to.default_content()

            # 点击侧边栏 (side_btn)
            # 你的源码显示 id="side_btn"
            try:
                side_btn = driver.find_element(By.ID, "side_btn")
                # 只有当它没展开时才点（简单判断一下位置，或者直接点也没事）
                driver.execute_script("arguments[0].click();", side_btn)
                time.sleep(1)
            except:
                pass

            # 点击跳过
            print("点击跳转下一任务...")
            try:
                # 你的源码显示 value="跳过&处理下一任务"
                skip_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@value, '跳过') and contains(@value, '下一任务')]")
                ))
                driver.execute_script("arguments[0].click();", skip_btn)
            except TimeoutException:
                print("❌ 找不到跳过按钮！")
                raise
            
            count += 1
            
            # --- 步骤6：检测弹窗 ---
            time.sleep(1.5) 
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                if "处理完" in alert_text or "列表" in alert_text:
                    print("\n" + "★"*30)
                    print("🎉 任务全部完成！")
                    print("★"*30)
                    alert.accept()
                    is_finished = True
                    break
                else:
                    alert.accept()
            except NoAlertPresentException:
                pass 

            # 等待新页面加载 (系统会自动跳转到新的长 URL)
            time.sleep(2)

        except UnexpectedAlertPresentException:
            try:
                driver.switch_to.alert.accept()
            except:
                pass
            continue

        except Exception as e:
            print(f"❌ 发生错误: {e}")
            time.sleep(5)
            
    print(f"\n程序退出。共保存 {count} 条数据。")
    input("按回车键关闭窗口...")

if __name__ == "__main__":
    main()
