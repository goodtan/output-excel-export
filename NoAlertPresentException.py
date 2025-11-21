import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoAlertPresentException, UnexpectedAlertPresentException, TimeoutException

# --- 配置 ---
EXCEL_FILE = '借贷人数据.xlsx'

def get_browser():
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-certificate-errors')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
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
    driver = get_browser()
    target_url = "http://10.200.18.179:8088/wcs/base/task/taskview.jsp" # 你的网址
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
            
            # 1. 定位数据行
            wait = WebDriverWait(driver, 10)
            borrower_tr = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//td[contains(text(), 'Borrower借贷人')]/..")
            ))

            # 2. 点击“显示”
            try:
                show_btn = borrower_tr.find_element(By.XPATH, ".//a[contains(text(), '显示')]")
                driver.execute_script("arguments[0].click();", show_btn)
                time.sleep(0.8) 
            except:
                pass 

            # 3. 提取数据
            cols = borrower_tr.find_elements(By.TAG_NAME, "td")
            
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

            # ==========================================
            # 4. 关键步骤：点击“催记”展开面板
            # ==========================================
            print("正在打开催记面板...")
            try:
                # 显式等待“催记”标签可点击
                # 这里的 XPath 匹配包含“催”和“记”文本的元素，通常是那个绿色竖条
                cuiji_tab = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(),'催') and contains(text(),'记')]")
                ))
                driver.execute_script("arguments[0].click();", cuiji_tab)
                
                # 【重要】等待1秒，让面板滑出来，否则下面的按钮可能点不到
                time.sleep(1) 
            except Exception as e:
                print(f"⚠️ 点击催记标签失败: {e}")
                # 如果失败，尝试继续找按钮，也许面板本来就是开着的

            # ==========================================
            # 5. 点击“跳过&处理下一任务”
            # ==========================================
            print("点击跳转下一任务...")
            # 等待跳过按钮出现并可点击
            skip_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[contains(@value, '跳过') and contains(@value, '下一任务')]")
            ))
            driver.execute_script("arguments[0].click();", skip_btn)
            
            count += 1
            
            # ==========================================
            # 6. 检测“完成”弹窗 (Alert)
            # ==========================================
            time.sleep(1) # 等待弹窗出现
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                print(f"检测到弹窗: {alert_text}")
                
                # 判断是否是结束语
                if "处理完" in alert_text or "列表" in alert_text:
                    print("\n" + "★"*30)
                    print("🎉 任务全部完成！脚本自动停止。")
                    print("★"*30)
                    alert.accept()
                    is_finished = True
                    break
                else:
                    # 其他无关弹窗，点确定忽略
                    alert.accept()
            except NoAlertPresentException:
                pass # 没弹窗，说明还有下一个任务

            # 等待页面完全加载下一条
            time.sleep(2)

        except UnexpectedAlertPresentException:
            # 如果在非预期的时候弹窗了，点掉它
            try:
                driver.switch_to.alert.accept()
            except:
                pass
            continue

        except Exception as e:
            print(f"❌ 发生错误: {e}")
            # 如果找不到按钮，可能是网络卡了，等待几秒重试
            time.sleep(3)
            
    print(f"\n程序退出。共保存 {count} 条数据。")

if __name__ == "__main__":
    main()
