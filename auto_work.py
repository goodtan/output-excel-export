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

def ensure_focus_on_latest_window(driver):
    """强制将焦点切换到最后一个窗口"""
    try:
        handles = driver.window_handles
        if not handles: return False
        driver.switch_to.window(handles[-1])
        return True
    except Exception:
        return False

def smart_switch_to_iframe(driver):
    """智能寻找包含数据的 iframe"""
    print("🔍 正在扫描数据框架...")
    driver.switch_to.default_content()
    
    # 1. 优先尝试 ID (frmcaseMainInfo)
    try:
        driver.switch_to.frame("frmcaseMainInfo")
        # 【关键】检查里面有没有我们需要的特定 ID 元素 phoneRole
        # 只有找到了 phoneRole，才说明我们真的进到了电话信息表格里
        if len(driver.find_elements(By.XPATH, "//*[contains(@id, 'phoneRole')]")) > 0:
            print("✅ 已切入数据框架 (frmcaseMainInfo)")
            return True
        driver.switch_to.default_content() 
    except:
        pass

    # 2. 暴力遍历
    iframe_list = driver.find_elements(By.TAG_NAME, "iframe")
    for i in range(len(iframe_list)):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(i)
            # 同样检查 phoneRole
            if len(driver.find_elements(By.XPATH, "//*[contains(@id, 'phoneRole')]")) > 0:
                print(f"✅ 在第 {i+1} 个框架中找到数据")
                return True
        except:
            continue
            
    print("❌ 当前页面未找到'电话信息'表格！")
    return False

def main():
    try:
        driver = get_browser()
    except Exception as e:
        print(f"❌ 浏览器启动失败: {e}")
        input("按回车键退出...")
        return

    driver.get(BASE_URL)

    print("\n" + "="*60)
    print("【请手动操作】")
    print("1. 登录系统。")
    print("2. 进入详情页（确保能看到电话信息表格）。")
    print("3. 回到这里，按【回车键】开始。")
    print("="*60 + "\n")
    input(">> 准备好后，按回车键开始: ") 

    ensure_focus_on_latest_window(driver)

    count = 0
    is_finished = False

    while not is_finished:
        try:
            print(f"\n>> 正在处理第 {count + 1} 个任务...")
            
            if not ensure_focus_on_latest_window(driver):
                break

            # 1. 找 iframe
            if not smart_switch_to_iframe(driver):
                print("⚠️ 无法定位数据，3秒后重试...")
                time.sleep(3)
                continue

            wait = WebDriverWait(driver, 10)
            
            # 2. 【精确修正】定位行
            # 只找包含 Borrower 且 ID 包含 phoneRole 的单元格
            try:
                borrower_tr = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//td[contains(text(), 'Borrower') and contains(@id, 'phoneRole')]/..")
                ))
            except TimeoutException:
                print("❌ 找不到 'Borrower借贷人' 行，跳过本条...")
                raise Exception("数据行未找到")

            # 3. 点击显示
            try:
                show_btn = borrower_tr.find_element(By.XPATH, ".//a[contains(text(), '显示')]")
                driver.execute_script("arguments[0].click();", show_btn)
                time.sleep(0.8) 
            except:
                pass 

            # 4. 【安全提取】提取数据
            cols = borrower_tr.find_elements(By.TAG_NAME, "td")
            
            if len(cols) >= 7:
                # 获取手机号
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
            else:
                print(f"⚠️ 警告：该行数据列数不足 ({len(cols)}列)，无法提取完整信息！")

            # 5. 切回主页面操作按钮
            driver.switch_to.default_content()

            # 点击侧边栏
            try:
                side_btn = driver.find_element(By.ID, "side_btn")
                driver.execute_script("arguments[0].click();", side_btn)
                time.sleep(1)
            except:
                pass

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
