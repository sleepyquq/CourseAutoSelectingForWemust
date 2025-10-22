from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

# 获取当前脚本所在目录的完整路径
current_dir = os.path.dirname(os.path.abspath(__file__))
# 构建msedgedriver.exe的完整路径
edge_driver_path = os.path.join(current_dir, 'msedgedriver.exe')

# 设置Edge选项
options = Options()
# 使用调试模式
options.add_argument("--remote-debugging-port=9222")
options.add_argument("--user-data-dir=C:\\Users\\ovo\\AppData\\Local\\Microsoft\\Edge\\User Data")
options.add_argument("--profile-directory=Default")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
options.add_argument("--disable-software-rasterizer")
options.add_experimental_option('detach', True)
options.add_experimental_option('excludeSwitches', ['enable-logging'])

try:
    # 启动Edge浏览器
    print("正在启动浏览器...")
    service = Service(edge_driver_path)
    driver = webdriver.Edge(service=service, options=options)

    # 最大化浏览器窗口
    driver.maximize_window()

    # 跳转到选课页面
    print("正在跳转到选课页面...")
    driver.get('https://eas-course-enrollment-wmweb.must.edu.mo/add-select')

    # 等待页面完全加载
    print("等待页面初始加载...")
    time.sleep(30)  # 增加初始加载等待时间

    # 打印当前URL，确认是否成功跳转
    print(f"当前页面URL: {driver.current_url}")

    while True:
        try:
            # 1. 点击选课分支
            print("\n尝试点击选课分支...")
            # 使用更精确的XPath定位
            branch_btn = driver.find_element(By.XPATH, "//li[contains(@class, 'ivu-menu-item')]//span[contains(text(), '人文藝術')]")
            print("找到选课分支按钮")
            branch_btn.click()
            print("已点击选课分支，等待页面响应...")
            time.sleep(5)  # 等待分支页面加载
            
            # 2. 点击加选按钮
            print("尝试点击加选按钮...")
            all_add_btns = driver.find_elements(By.XPATH, "//span[@class='text-button' and contains(text(), '加選')]")
            print(f"当前页面找到 {len(all_add_btns)} 个加選按钮")
            for idx, btn in enumerate(all_add_btns):
                print(f"第{idx+1}个加選按钮HTML: {btn.get_attribute('outerHTML')}")

            if all_add_btns:
                select_btn = all_add_btns[2]
                driver.execute_script("arguments[0].scrollIntoView(true);", select_btn)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", select_btn)
                print("已点击第一个加选按钮，等待弹窗...")
                time.sleep(5)
            else:
                print("未找到加選按钮，跳过本轮。")
            
            # 3. 刷新页面，准备下一次尝试
            print("准备刷新页面...")
            driver.refresh()
            print("页面已刷新，等待重新加载...")
            time.sleep(10)  # 等待页面刷新完成

        except Exception as e:
            print(f"\n操作失败：{str(e)}")
            print("刷新页面重试...")
            driver.refresh()
            print("等待页面重新加载...")
            time.sleep(15)  # 失败后等待15秒

except Exception as e:
    print(f"浏览器启动失败：{str(e)}")
    print("\n请按以下步骤操作：")
    print("1. 关闭所有Edge浏览器窗口")
    print("2. 打开任务管理器，结束所有msedge.exe进程")
    print("3. 以管理员身份运行PowerShell或命令提示符")
    print("4. 重新运行脚本")