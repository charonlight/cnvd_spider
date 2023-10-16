"""更新于 2022-12-14"""
import os.path
import time
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver import Chrome


# 获取驱动
def get_webdriver():
    opt = init_param()  # 获取webdriver参数设置

    # # 方式一 自动更新驱动
    # base_dir = os.path.dirname(os.path.abspath(__file__))  # 该py文件所在的目录
    # driver_path = ChromeDriverManager(path=base_dir).install()  # 下载驱动并返回驱动路径
    # driver = webdriver.Chrome(driver_path, opt)  # 创建浏览器对象

    # pip install webdriver-manager==4.0.1
    driver_path = ChromeDriverManager().install()  # 下载驱动并返回驱动路径
    driver = webdriver.Chrome(executable_path=driver_path, chrome_options=opt)  # 创建浏览器对象

    # # 方式二 需手动下载驱动
    # # 驱动下载地址: http://chromedriver.storage.googleapis.com/index.html
    # driver = Chrome(chrome_options=opt)

    # 打开浏览器之前执行指定路径下的js代码
    # execute_js()

    # 设置等待等全局参数
    driver.maximize_window()  # 打开最大窗口
    driver.implicitly_wait(5)  # 设置等待时间

    return driver


# 初始化webdriver参数
def init_param():
    opt = Options()  # chrome_options 初始化选项

    # 设置浏览器初始 位置x,y & 宽高x,y
    # opt.add_argument(f'--window-position={217},{172}')
    # opt.add_argument(f'--window-size={1200},{1000}')

    # 关闭自动测试状态显示 // 会导致浏览器报：请停用开发者模式 ###window.navigator.webdriver还是返回True,当返回undefined时应该才可行。
    opt.add_experimental_option('excludeSwitches', ['enable-automation'])  # 开启实验性功能

    # 关闭开发者模式
    # opt.add_experimental_option("useAutomationExtension", False)
    opt.add_argument('--disable-blink-features=AutomationControlled')  # 关闭开发者模式 88版本之后

    # 禁止图片加载
    # prefs = {"profile.managed_default_content_settings.images": 2}
    # opt.add_experimental_option("prefs", prefs)

    # 设置中文
    # opt.add_argument('lang=zh_CN.UTF-8')

    # 更换头部
    # opt.add_argument('user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36"')

    # 无头浏览器
    opt.add_argument('--headless')  # 隐藏浏览器

    # 部署项目在linux时，其驱动会要求这个参数
    # opt.add_argument('--no-sandbox')

    return opt


# 打开浏览器之前执行指定路径下的js代码
def execute_js(driver):
    # 打开浏览器之前执行指定路径下的js代码
    with open('./config/stealth.min.js') as f:
        js = f.read()
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": js})
    # driver.get('https://bot.sannysoft.com/')
    time.sleep(1)
    driver.save_screenshot('walkaround1.png')

    # # 执行固定的js代码
    # driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    #     "source": """
    #   navigator.webdriver = undefined
    #     Object.defineProperty(navigator, 'webdriver', {
    #       get: () => undefined
    #     })
    #   """
    # })
    # # 一样的js
    # driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    #     "source": """Object.defineProperty(navigator, 'webdriver', {get: () => undefined})""",
    # })


if __name__ == '__main__':
    web = get_webdriver()
    web.get('https://www.baidu.com/')
    print(web.title)
