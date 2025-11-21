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
    """
    【核心黑科技】智能寻找包含数据的 iframe
    不依赖 ID，而是挨个进去看有没有'Borrower借贷人'这几个字
    """
    print("🔍 正在扫描页面框架，寻找数据...")
    
    # 1. 先切回最外层
    driver.switch_to.default_content()
    
    # 检查最外层有没有
    if len(driver.find_elements(By.XPATH, "//*[contains(text(), 'Borrower借贷人')]")) > 0:
        print("✅ 数据就在最外层，无需切换")
        return True

    # 2. 尝试根据 ID 直接切（为了兼容性保留这步）
    try:
        driver.switch_to.frame("frmcaseMainInfo")
        # 检查切进去对不对
        if len(driver.find_elements(By.XPATH, "//*[contains(text(), 'Borrower借贷人')]")) > 0:
            print("✅ 通过 ID 锁定数据框架")
            return True
        driver.switch_to.default_content() # 不对就退出来
    except:
        pass

    # 3. 暴力遍历所有 iframe
    iframe_list = driver.find_elements(By.TAG_NAME, "iframe")
    print(f"ℹ️ 发现 {len(iframe_list)} 个潜在框架，正在逐一排查...")
    
    for i in range(len(iframe_list)):
        try:
            driver.switch_to.default_content() # 每次都从头开始
            driver.switch_to.frame(i) # 切入第 i 个框架
            
            # 看看有没有我们要找的文字
            # 这里用 find_elements 如果找不到不会报错只会返回空列表
            elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Borrower借贷人')]")
            
            if len(elements) > 0:
                print(f"✅ 成功在第 {i+1} 个框架中找到数据！")
                return True
                
        except Exception as e:
            print(f"⚠️ 扫描框架 {i} 出错: {e}")
            continue
            
    # 如果都找不到，尝试能不能打印源代码看看
    print("❌ 扫描结束，未找到包含数据的框架！")
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
    print("请手动登录系统...")
    print("进入第一个任务详情页后，在控制台按【回车键】开始")
    print("="*50 + "\n")
    input() 

    count = 0
    is_finished = False

    while not is_finished:
        try:
            print(f"\n>> 正在处理第 {count + 1} 个任务...")
            
            # --- 步骤1：智能寻找并切入 iframe ---
            if not smart_switch_to_iframe(driver):
                print("❌ 无法定位数据区域。请确认页面已加载完成，且处于详情页。")
                # 暂停一下让用户看清
                time.sleep(3)
                # 尝试刷新或者跳过
                raise Exception("Iframe定位失败")

            # --- 步骤2：定位并操作数据 ---
            wait = WebDriverWait(driver, 10)
            
            # 定位行
            borrower_tr = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//td[contains(text(), 'Borrower')]/..")
            ))

            # 点击显示
            try:
                show_btn = borrower_tr.find_element(By.XPATH, ".//a[contains(text(), '显示')]")
                driver.execute_script("arguments[0].click();", show_btn)
                time.sleep(0.8) 
            except:
                pass 

            # 提取数据
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

            # --- 步骤3：切回主页面操作按钮 ---
            print("切换回主页面操作按钮...")
            driver.switch_to.default_content()

            # 点击侧边栏 (side_btn)
            try:
                side_btn = driver.find_element(By.ID, "side_btn")
                driver.execute_script("arguments[0].click();", side_btn)
                time.sleep(1)
            except:
                # 备用方案：如果ID找不到，用XPath找绿色块
                try:
                    side_btn = driver.find_element(By.XPATH, "//a[contains(@class, 'side_btn')]")
                    driver.execute_script("arguments[0].click();", side_btn)
                    time.sleep(1)
                except:
                    print("⚠️ 无法点击侧边栏，尝试直接点击跳过")

            # 点击跳过
            print("点击跳转下一任务...")
            try:
                skip_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@value, '跳过') and contains(@value, '下一任务')]")
                ))
                driver.execute_script("arguments[0].click();", skip_btn)
            except TimeoutException:
                print("❌ 找不到跳过按钮！")
                raise
            
            count += 1
            
            # --- 步骤4：检测弹窗 ---
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
