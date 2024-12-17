from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service

# 启动Edge浏览器
driver = webdriver.Edge(service=Service(executable_path='D:\\软件\\pycharm-community-2023.1.4\\pythonProject1\\edgedriver_win64\\msedgedriver.exe'))

# 导航到目标网页
driver.get('https://lishi.tianqi.com/qingdao/202410.html')  # 确保URL是正确的

try:
    # 等待"查看更多"的<div>元素可见并点击
    # 这里假设<div>元素可以通过类名"lishidesc2"定位
    more_data_div = WebDriverWait(driver, 1000).until(
        EC.visibility_of_element_located((By.CLASS_NAME, "lishidesc2"))
    )

    more_data_div.click()
except TimeoutException:
    print("未在规定时间内出现")

# 等待更多数据加载完成
# ...（这里可能需要根据实际情况添加等待更多数据加载的逻辑）

# 关闭浏览器
driver.quit()