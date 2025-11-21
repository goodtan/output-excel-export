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

def get_browser():
    """启动 Google Chrome 浏览器 (离线版)"""
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-certificate-errors')
    # 禁用自动化控制条，防止被检测
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    # 获取当前目录下的 chromedriver.exe
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
        
    driver_path = os.path.join(base_path, "chromedriver.exe")
    print(f"正在加载驱动: {driver_path}")

    if not os.path.exists(driver_path):
        print("❌ 错误：找不到 chromedriver.exe，请确保它和脚本在同一目录下！")
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

def main():
    try:
        driver = get_browser()
    except Exception as e:
        print(f"❌ 浏览器启动失败: {e}")
        input("按回车键退出...")
        return

    # 你的登录地址
    target_url = "http://10.200.18.179:8088/wcs/base/task/taskview.jsp"
    driver.get(target_url)

    print("\n" + "="*50)
    print("请手动登录系统...")
    print("进入第一个任务详情页后，在控制台按【回车键】开始")
    print("="*50 + "\n")
    input() 

    count = 0
    is_finished = False

    while not is_finished:
        try:
            print(f"\n>> 正在处理第 {count + 1} 个任务...")
            
            # ==========================================
            # 步骤 1: 进入 iframe 抓取数据
            # ==========================================
            # 你的HTML显示iframe的ID是 frmcaseMainInfo，我们直接切进去
            driver.switch_to.default_content() # 确保在最外层
            
            wait = WebDriverWait(driver, 10)
            try:
                # 显式等待 iframe 出现并切入
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frmcaseMainInfo")))
            except TimeoutException:
                print("❌ 找不到数据框架 (frmcaseMainInfo)，请确认是否在详情页？")
                raise

            # 定位 'Borrower借贷人' 行
            try:
                borrower_tr = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//td[contains(text(), 'Borrower')]/..")
                ))
            except TimeoutException:
                print("❌ 找不到 'Borrower借贷人' 行，跳过此步骤...")
                # 这里如果不抛出异常，下面的代码会报错，所以抛出重试
                raise Exception("数据行未找到")

            # 点击“显示”
            try:
                show_btn = borrower_tr.find_element(By.XPATH, ".//a[contains(text(), '显示')]")
                driver.execute_script("arguments[0].click();", show_btn)
                time.sleep(0.8) 
            except:
                pass 

            # 提取数据
            cols = borrower_tr.find_elements(By.TAG_NAME, "td")
            # 获取手机号 (id包含 phoneValue)
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

            # ==========================================
            # 步骤 2: 切回主页面 点击按钮
            # ==========================================
            # 【关键】按钮在 iframe 外面，必须切出来！
            driver.switch_to.default_content()

            # 点击侧边栏开关 (id="side_btn")
            print("正在打开催记面板...")
            try:
                # 根据HTML，按钮ID是 side_btn
                side_btn = wait.until(EC.element_to_be_clickable((By.ID, "side_btn")))
                driver.execute_script("arguments[0].click();", side_btn)
                time.sleep(1) # 等待滑出动画
            except Exception as e:
                print(f"⚠️ 点击侧边栏失败: {e}")

            # 点击跳过按钮
            # 根据HTML，value="跳过&处理下一任务"
            print("点击跳转下一任务...")
            try:
                skip_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@value, '跳过') and contains(@value, '下一任务')]")
                ))
                driver.execute_script("arguments[0].click();", skip_btn)
            except TimeoutException:
                print("❌ 找不到跳过按钮！请确认侧边栏是否已展开？")
                raise
            
            count += 1
            
            # ==========================================
            # 步骤 3: 检测完成弹窗
            # ==========================================
            time.sleep(1.5) 
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                # 包含 "处理完" 或 "列表" 即视为结束
                if "处理完" in alert_text or "列表" in alert_text:
                    print("\n" + "★"*30)
                    print(f"🎉 检测到弹窗: [{alert_text}]")
                    print("🎉 任务全部完成！")
                    print("★"*30)
                    alert.accept()
                    is_finished = True
                    break
                else:
                    alert.accept()
            except NoAlertPresentException:
                pass 

            # 等待新页面加载
            time.sleep(2)

        except UnexpectedAlertPresentException:
            try:
                driver.switch_to.alert.accept()
            except:
                pass
            continue

        except Exception as e:
            print(f"❌ 发生错误: {e}")
            time.sleep(3)
            
    print(f"\n程序退出。共保存 {count} 条数据。")
    input("按回车键关闭窗口...")

if __name__ == "__main__":
    main()
