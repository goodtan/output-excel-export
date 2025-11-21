import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import NoAlertPresentException, UnexpectedAlertPresentException, TimeoutException

# --- 配置 ---
EXCEL_FILE = '借贷人数据.xlsx'

def get_browser():
    """启动 Edge 浏览器"""
    options = webdriver.EdgeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-certificate-errors')
    driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
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
    【核心修复】自动查找并切换到包含数据的 iframe
    """
    try:
        # 1. 先回到最外层，防止重复切换报错
        driver.switch_to.default_content()
        
        # 2. 查找页面里所有的 iframe
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        
        if len(iframes) > 0:
            print(f"🔍 检测到页面有 {len(iframes)} 个 iframe，尝试切换到第一个...")
            # 通常这种系统的主体内容都在第一个或第二个 iframe 里
            # 这里尝试切换到第一个可见的 iframe
            driver.switch_to.frame(0) 
            # 如果你的系统很复杂，可能需要改成 driver.switch_to.frame("mainFrame") 或其他ID
            print("已进入 iframe 内部")
        else:
            print("ℹ️ 未检测到 iframe，在主页面查找")
            
    except Exception as e:
        print(f"⚠️ 切换 iframe 失败 (非致命错误): {e}")

def main():
    driver = get_browser()
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
            
            # --- 【修复步骤】每次循环开始前，先尝试切入 iframe ---
            switch_to_content_iframe(driver)

            # 1. 定位数据行 (增加超时重试)
            wait = WebDriverWait(driver, 10)
            try:
                # 尝试定位 'Borrower'，这里用模糊匹配 'Borrower' 防止空格问题
                borrower_tr = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//td[contains(text(), 'Borrower')]/..")
                ))
            except TimeoutException:
                # 如果切了 iframe 还是找不到，打印当前页面源码的一小部分帮我分析
                print("❌ 找不到 'Borrower借贷人' 行！")
                print("可能原因：1. 没在这个 iframe 里  2. 页面没加载出来")
                print("当前页面HTML片段:", driver.page_source[:500]) # 打印前500字符看看是不是还在登录页
                raise Exception("元素定位超时")

            # 2. 点击“显示”
            try:
                show_btn = borrower_tr.find_element(By.XPATH, ".//a[contains(text(), '显示')]")
                driver.execute_script("arguments[0].click();", show_btn)
                time.sleep(1) # 稍微多等一点点
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
                # 既然都在 iframe 里了，这里直接找
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
            
            # 6. 检测弹窗 (检测弹窗时，不需要管 iframe，Selenium 会自动处理 Alert)
            time.sleep(1.5) 
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                print(f"检测到弹窗: {alert_text}")
                
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
            # 出错后等待时间加长，方便你看清屏幕
            time.sleep(5)
            # 询问是否重试
            # break # 如果你想出错就停止，取消这行注释
            
    print(f"\n程序退出。共保存 {count} 条数据。")
    input("按回车键关闭窗口...")

if __name__ == "__main__":
    main()
