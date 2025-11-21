import time
import pandas as pd
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# 【修改点1】引入 Chrome 的服务
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoAlertPresentException, UnexpectedAlertPresentException, TimeoutException

# --- 配置 ---
EXCEL_FILE = '借贷人数据.xlsx'

def get_browser():
    """启动 Google Chrome 浏览器 (离线版)"""
    # 【修改点2】使用 ChromeOptions
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized') # 最大化
    options.add_argument('--ignore-certificate-errors') # 忽略证书错误
    
    # 禁止浏览器显示“正在受到自动测试软件的控制”提示（可选优化）
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    # 【修改点3】手动寻找本地的 chromedriver.exe
    if getattr(sys, 'frozen', False):
        # 如果是打包后的 exe
        base_path = os.path.dirname(sys.executable)
    else:
        # 如果是脚本运行
        base_path = os.path.dirname(os.path.abspath(__file__))
        
    # 指定驱动文件名
    driver_path = os.path.join(base_path, "chromedriver.exe")

    print(f"正在加载驱动: {driver_path}")

    if not os.path.exists(driver_path):
        print("\n" + "!"*50)
        print("❌ 错误：找不到 chromedriver.exe")
        print(f"请确保 'chromedriver.exe' 文件位于文件夹: {base_path}")
        print("!"*50 + "\n")
        input("按回车键退出...")
        sys.exit(1)

    # 启动 Chrome
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

def switch_to_content_iframe(driver):
    """自动查找并切换到包含数据的 iframe"""
    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if len(iframes) > 0:
            # 尝试切入第一个 iframe
            driver.switch_to.frame(0) 
        else:
            pass
    except Exception as e:
        print(f"⚠️ 切换 iframe 失败: {e}")

def main():
    try:
        driver = get_browser()
    except Exception as e:
        print(f"❌ 浏览器启动失败: {e}")
        print("可能原因：驱动版本与浏览器版本不匹配，或驱动文件缺失。")
        input("按回车键退出...")
        return

    # 你的系统地址
    target_url = "http://10.200.18.179:8088/wcs/base/task/taskview.jsp"
    driver.get(target_url)

    print("\n" + "="*50)
    print("已启动 Google Chrome...")
    print("请手动登录系统...")
    print("进入第一个任务详情页后，在控制台按【回车键】开始")
    print("="*50 + "\n")
    input() 

    count = 0
    is_finished = False

    while not is_finished:
        try:
            print(f"\n>> 正在处理第 {count + 1} 个任务...")
            
            # 尝试切入 iframe
            switch_to_content_iframe(driver)

            wait = WebDriverWait(driver, 10)
            
            # 1. 定位借贷人行
            try:
                borrower_tr = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//td[contains(text(), 'Borrower')]/..")
                ))
            except TimeoutException:
                print("❌ 找不到 'Borrower借贷人' 行！")
                raise Exception("元素定位超时")

            # 2. 点击显示
            try:
                show_btn = borrower_tr.find_element(By.XPATH, ".//a[contains(text(), '显示')]")
                driver.execute_script("arguments[0].click();", show_btn)
                time.sleep(1) 
            except:
                pass 

            # 3. 提取数据
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

            # 4. 点击催记
            print("正在打开催记面板...")
            try:
                cuiji_tab = driver.find_element(By.XPATH, "//*[contains(text(),'催') and contains(text(),'记')]")
                driver.execute_script("arguments[0].click();", cuiji_tab)
                time.sleep(1) 
            except Exception as e:
                print(f"⚠️ 点击催记失败: {e}")

            # 5. 点击跳过
            print("点击跳转下一任务...")
            skip_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[contains(@value, '跳过') and contains(@value, '下一任务')]")
            ))
            driver.execute_script("arguments[0].click();", skip_btn)
            
            count += 1
            
            # 6. 检测弹窗
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
