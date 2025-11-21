import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
# 引入处理弹窗的异常类
from selenium.common.exceptions import NoAlertPresentException, UnexpectedAlertPresentException

# --- 配置 ---
EXCEL_FILE = '借贷人数据.xlsx'  # 结果保存的文件名

def get_browser():
    """启动浏览器配置"""
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-certificate-errors')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def save_to_excel(data_dict):
    """保存数据到Excel (追加模式，实时保存)"""
    df_new = pd.DataFrame([data_dict])
    if not os.path.exists(EXCEL_FILE):
        df_new.to_excel(EXCEL_FILE, index=False)
    else:
        df_old = pd.read_excel(EXCEL_FILE)
        df_all = pd.concat([df_old, df_new], ignore_index=True)
        df_all.to_excel(EXCEL_FILE, index=False)
    print(f"✅ 数据已写入Excel: {data_dict.get('姓名')}")

def main():
    driver = get_browser()
    
    # 填入你的登录网址
    target_url = "http://10.200.18.179:8088/wcs/base/task/taskview.jsp" 
    driver.get(target_url)

    print("\n" + "="*50)
    print("请手动登录...")
    print("进入第一个任务详情页后，在控制台按【回车键】开始")
    print("="*50 + "\n")
    input() 

    count = 0
    is_finished = False # 标记是否完成

    while not is_finished:
        try:
            print(f"\n>> 正在处理第 {count + 1} 个任务...")
            
            # 1. 等待数据加载 (寻找 Borrower借贷人)
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
                pass # 可能已经显示了

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
            
            # 4. 保存数据
            save_to_excel(data)

            # 5. 点击催记 (展开侧边栏)
            try:
                cuiji_tab = driver.find_element(By.XPATH, "//*[contains(text(),'催') and contains(text(),'记')]")
                driver.execute_script("arguments[0].click();", cuiji_tab)
                time.sleep(0.5)
            except:
                pass

            # 6. 点击“跳过&处理下一任务”
            print("点击跳转下一任务...")
            skip_btn = driver.find_element(By.XPATH, "//input[@value='跳过&处理下一任务']")
            driver.execute_script("arguments[0].click();", skip_btn)
            
            count += 1
            
            # --- 【核心修改】检测是否弹出“任务已处理完” ---
            time.sleep(1) # 稍微等待弹窗出现
            try:
                # 切换到弹窗上下文
                alert = driver.switch_to.alert
                alert_text = alert.text
                print(f"检测到弹窗内容: {alert_text}")
                
                # 判断弹窗文字是否包含关键信息
                if "处理完" in alert_text or "列表" in alert_text:
                    print("\n" + "★"*30)
                    print("🎉 所有任务已处理完毕！")
                    print("★"*30)
                    alert.accept() # 点击弹窗的“确定”
                    is_finished = True # 结束循环标记
                    break # 跳出循环
                else:
                    # 如果是其他弹窗（比如报错），点击确定继续运行
                    alert.accept()
            except NoAlertPresentException:
                # 如果没有弹窗，说明还有任务，继续循环
                pass

            # 等待页面刷新进入下一条
            time.sleep(2)

        except UnexpectedAlertPresentException:
            # 处理意外弹窗的情况
            try:
                driver.switch_to.alert.accept()
            except:
                pass
            continue

        except Exception as e:
            print(f"❌ 发生错误: {e}")
            # 如果页面卡住，可以选择手动干预，这里设置等待
            time.sleep(3)
            # 如果连续报错，可以选择 break

    print(f"\n程序结束。共抓取 {count} 条数据。")
    print(f"Excel文件位置: {os.path.abspath(EXCEL_FILE)}")
    # driver.quit() # 如果你想保留浏览器查看，注释掉这行

if __name__ == "__main__":
    main()
