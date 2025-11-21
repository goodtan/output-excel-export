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
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    # 手动寻找本地的 chromedriver.exe
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
        
    driver_path = os.path.join(base_path, "chromedriver.exe")
    print(f"正在加载驱动: {driver_path}")

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

def switch_to_content_iframe(driver):
    """
    【核心修复】精准切换到 ID 为 frmcaseMainInfo 的 iframe
    """
    try:
        # 1. 先回到最外层
        driver.switch_to.default_content()
        
        # 2. 等待并切换到指定的 iframe
        # 你的截图显示 iframe 的 id="frmcaseMainInfo"
        wait = WebDriverWait(driver, 5) # 给5秒时间找这个框架
        
        print("🔍 正在寻找并切换到数据框架 (frmcaseMainInfo)...")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frmcaseMainInfo")))
        
        print("✅ 已成功进入数据框架内部")
        return True

    except TimeoutException:
        print("⚠️ 警告：找不到 ID 为 'frmcaseMainInfo' 的框架！")
        # 备用方案：如果是嵌套结构，有时候需要先切父框架再切子框架
        # 但通常 ID 定位是最准的。如果这里报错，可能在错误的页面。
        return False
    except Exception as e:
        print(f"⚠️ 切换框架失败: {e}")
        return False

def main():
    try:
        driver = get_browser()
    except Exception as e:
        print(f"❌ 浏览器启动失败: {e}")
        input("按回车键退出...")
        return

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
            
            # --- 步骤1：切入 iframe ---
            if not switch_to_content_iframe(driver):
                print("❌ 无法进入数据区域，脚本暂停。请确认你是否在任务详情页？")
                # 可以在这里加个 input 暂停，方便你手动调整页面
                # input("调整好页面后按回车继续...") 
                # continue 
            
            wait = WebDriverWait(driver, 10)
            
            # --- 步骤2：定位数据行 ---
            try:
                # 截图显示 td 的 id 包含 phoneRole，我们也可以利用这个特征
                # 或者继续用文本定位，这里加了重试机制
                borrower_tr = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//td[contains(text(), 'Borrower') or contains(@id, 'phoneRole')]/..")
                ))
            except TimeoutException:
                print("❌ 在当前框架内找不到 'Borrower借贷人' 行！")
                # 打印一点页面源码来看看是不是切错了
                # print(driver.page_source[:500])
                raise Exception("元素定位超时")

            # --- 步骤3：点击显示 ---
            try:
                show_btn = borrower_tr.find_element(By.XPATH, ".//a[contains(text(), '显示')]")
                driver.execute_script("arguments[0].click();", show_btn)
                time.sleep(1) 
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

            # --- 步骤5：切回主文档点击跳过 ---
            # 注意！“催记”侧边栏和“跳过”按钮通常在 iframe 外面（主页面）！
            # 所以提取完数据后，必须切出来
            print("正在切换回主页面进行操作...")
            driver.switch_to.default_content()

            # 点击催记
            try:
                cuiji_tab = driver.find_element(By.XPATH, "//*[contains(text(),'催') and contains(text(),'记')]")
                driver.execute_script("arguments[0].click();", cuiji_tab)
                time.sleep(1) 
            except:
                pass # 可能不需要点，或者已经在外面了

            # 点击跳过
            print("点击跳转下一任务...")
            try:
                skip_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@value, '跳过') and contains(@value, '下一任务')]")
                ))
                driver.execute_script("arguments[0].click();", skip_btn)
            except TimeoutException:
                print("❌ 找不到跳过按钮！请确认催记面板是否展开？")
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
