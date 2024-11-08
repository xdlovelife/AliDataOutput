import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import zipfile
import random
from selenium.webdriver.common.keys import Keys
import pandas as pd
import xlwt
import win32com.client
import psutil

# 配置文件路径
CONFIG_FILE = 'app_config.json'


class Logger:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def log(self, message, level="INFO"):
        timestamp = time.strftime("[%Y-%m-%d %H:%M:%S]")
        log_message = f"{timestamp} [{level}] {message}\n"
        self.text_widget.insert(tk.END, log_message)
        self.text_widget.see(tk.END)
        self.text_widget.update()


def save_config(excel_path, account, password, driver_path):
    """保存所有配置信息"""
    config = {
        'excel_path': excel_path,
        'account': account,
        'password': password,
        'driver_path': driver_path
    }
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False)


def load_config():
    """加载配置信息"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}


def check_local_chromedriver(logger):
    """检查本地是否存在ChromeDriver"""
    common_paths = [
        "D:\\chromedriver-win64\\chromedriver.exe",
        "D:\\chromedriver_win32\\chromedriver.exe",
        "D:\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe",
        "D:\\chromedriver_win32\\chromedriver-win32\\chromedriver.exe"
    ]

    for path in common_paths:
        if os.path.exists(path):
            logger.log(f"找到本地ChromeDriver: {path}", "SUCCESS")
            return path

    logger.log("未找到本地ChromeDriver", "INFO")
    return None


def get_chrome_version():
    """获取Chrome浏览器版本"""
    chrome_path = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
    if os.path.exists(chrome_path):
        from win32com.client import Dispatch
        parser = Dispatch("Scripting.FileSystemObject")
        try:
            version = parser.GetFileVersion(chrome_path)
            return version.split('.')[0]
        except Exception:
            return None
    return None


def check_internet_connection(logger):
    """检查网络连接"""
    sites = [
        "https://www.alibaba.com",
        "https://www.taobao.com",
        "https://www.baidu.com"
    ]

    for site in sites:
        try:
            response = requests.get(site, timeout=10)
            if response.status_code == 200:
                logger.log(f"网络连接正常 ({site})", "SUCCESS")
                return True
        except:
            continue

    logger.log("网络连接失败，请检查网络设置", "ERROR")
    return False


def get_chrome_config():
    """获取Chrome配置"""
    options = webdriver.ChromeOptions()

    # 基本设置
    options.add_argument('--no-sandbox')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-software-rasterizer')

    # 性能优化
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-logging')
    options.add_argument('--disable-notifications')
    options.add_argument('--disable-default-apps')

    # 避免检测
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)

    # 设置窗口大小
    options.add_argument('--window-size=1000,800')

    return options


def download_chromedriver_from_official(logger):
    """从Chrome官方下载ChromeDriver"""
    try:
        chrome_version = get_chrome_version()
        if not chrome_version:
            logger.log("无法获取Chrome版本", "ERROR")
            return None

        logger.log(f"检测到Chrome版本: {chrome_version}", "INFO")

        # 准备目录
        driver_dir = os.path.join(os.getcwd(), "chromedriver")
        os.makedirs(driver_dir, exist_ok=True)
        driver_path = os.path.join(driver_dir, "chromedriver.exe")

        # 如果已存在驱动，直接返回
        if os.path.exists(driver_path):
            logger.log(f"使用已存在的ChromeDriver: {driver_path}", "INFO")
            return driver_path

        # 获取具体版本号
        logger.log("正在获取最新版本信息...", "INFO")
        version_url = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_" + chrome_version

        try:
            response = requests.get(version_url, timeout=10)
            if not response.ok:
                logger.log("无法获取版本信息", "ERROR")
                return None

            specific_version = response.text.strip()
            logger.log(f"找到匹配版本: {specific_version}", "INFO")

            # 构建下载URL
            download_url = f"https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/{specific_version}/win64/chromedriver-win64.zip"
            logger.log(f"正在从官方源下载: {download_url}", "INFO")

            # 下载文件
            response = requests.get(download_url, stream=True, timeout=30)
            if not response.ok:
                logger.log("下载失败", "ERROR")
                return None

            # 保存文件
            zip_path = os.path.join(driver_dir, "chromedriver.zip")
            total_size = int(response.headers.get('content-length', 0))
            block_size = 1024

            with open(zip_path, 'wb') as f:
                downloaded = 0
                for data in response.iter_content(block_size):
                    downloaded += len(data)
                    f.write(data)
                    if total_size > 0:
                        percent = int((downloaded / total_size) * 100)
                        if percent % 10 == 0:
                            logger.log(f"下载进度: {percent}%", "INFO")

            logger.log("下载完成，正在解压...", "INFO")

            # 解压文件
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(driver_dir)

            # 移动文件到正确位置
            chromedriver_dir = os.path.join(driver_dir, "chromedriver-win64")
            if os.path.exists(chromedriver_dir):
                source_driver = os.path.join(chromedriver_dir, "chromedriver.exe")
                if os.path.exists(source_driver):
                    import shutil
                    shutil.move(source_driver, driver_path)
                    shutil.rmtree(chromedriver_dir)

            # 删除zip文件
            os.remove(zip_path)

            if os.path.exists(driver_path):
                logger.log(f"ChromeDriver成功: {driver_path}", "SUCCESS")
                return driver_path
            else:
                logger.log("ChromeDriver安装失败", "ERROR")
                return None

        except Exception as e:
            logger.log(f"载过程中发生错误: {str(e)}", "ERROR")
            return None

    except Exception as e:
        logger.log(f"载ChromeDriver时发生错误: {str(e)}", "ERROR")
        return None


def get_manual_driver_path(logger):
    """动择ChromeDriver路径"""
    logger.log("请手动选择ChromeDriver文件...", "INFO")
    driver_path = filedialog.askopenfilename(
        title='选择ChromeDriver文件',
        filetypes=[('ChromeDriver', 'chromedriver.exe'), ('All Files', '*.*')]
    )

    if driver_path:
        logger.log(f"已选择ChromeDriver: {driver_path}", "SUCCESS")
        return driver_path
    return None


def get_driver_path(logger):
    """获取ChromeDriver路径"""
    # 首先检查配置文件
    config = load_config()
    driver_path = config.get('driver_path')
    if driver_path and os.path.exists(driver_path):
        logger.log(f"使用已保存的ChromeDriver配置: {driver_path}", "INFO")
        return driver_path

    # 然后检查本地目录
    driver_path = check_local_chromedriver(logger)
    if driver_path:
        # 保存找到的路径到配置文件
        config['driver_path'] = driver_path
        save_config(
            config.get('excel_path', ''),
            config.get('account', ''),
            config.get('password', ''),
            driver_path
        )
        return driver_path

    # 如果本地没有，尝试自动载
    logger.log("尝试自动下载ChromeDriver...", "INFO")
    driver_path = download_chromedriver_from_official(logger)

    # 如果自动下载失败，提示手动选择
    if not driver_path:
        logger.log("自动下载失败，请手动选择ChromeDriver文件", "WARNING")
        result = messagebox.askyesno("提示",
                                     "未找到ChromeDriver，是否手动选择ChromeDriver文件？\n\n" +
                                     "您可以从以下地址下载对应版本的ChromeDriver：\n" +
                                     "https://googlechromelabs.github.io/chrome-for-testing/\n\n" +
                                     "请确保下载的版本与Chrome浏览器版本匹配。"
                                     )

        if result:
            driver_path = get_manual_driver_path(logger)
            if driver_path:
                # 保存成功选择的路径
                config['driver_path'] = driver_path
                save_config(
                    config.get('excel_path', ''),
                    config.get('account', ''),
                    config.get('password', ''),
                    driver_path
                )
            else:
                logger.log("未选择ChromeDriver文件，操作取消", "WARNING")
                return None
        else:
            logger.log("操作取消", "WARNING")
            return None

    return driver_path


class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("阿里巴巴数据处理工具")
        self.root.geometry("800x600")

        # 设置窗口图标
        try:
            # 检查图标文件是否存在
            icon_path = 'xdlovelife.ico'
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                # 如果图标文件不存在，尝试从资源目录加载
                resource_icon_path = os.path.join('resources', 'xdlovelife.ico')
                if os.path.exists(resource_icon_path):
                    self.root.iconbitmap(resource_icon_path)
                else:
                    self.logger.log("未找到图标文件: xdlovelife.ico", "WARNING")
        except Exception as e:
            self.logger.log(f"设置图标时发生错: {str(e)}", "WARNING")

        # 创建界面
        self.create_widgets()

        # 加载保存的配置
        self.load_saved_config()

    def load_saved_config(self):
        """加载保存的配置并填充到界面"""
        config = load_config()
        if config:
            # 填Excel路径
            if 'excel_path' in config and os.path.exists(config['excel_path']):
                self.excel_path.set(config['excel_path'])

            # 填充账号
            if 'account' in config:
                self.account_entry.delete(0, tk.END)
                self.account_entry.insert(0, config['account'])

            # 填充密码
            if 'password' in config:
                self.password_entry.delete(0, tk.END)
                self.password_entry.insert(0, config['password'])

    def select_excel(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title='选择Excel文件',
            filetypes=[('Excel Files', '*.xlsx;*.xls'), ('All Files', '*.*')]
        )
        if file_path:
            self.excel_path.set(file_path)
            # 保存配置
            self.save_current_config()

    def save_current_config(self):
        """保存当前配置"""
        config = load_config()
        save_config(
            self.excel_path.get(),
            self.account_entry.get(),
            self.password_entry.get(),
            config.get('driver_path', '')
        )

    def execute(self):
        """执行操作"""
        excel_path = self.excel_path.get()
        account = self.account_entry.get()
        password = self.password_entry.get()

        # 保存当前配置
        self.save_current_config()

        # 执行操作
        execute_action(excel_path, account, password, self.logger)

    def create_widgets(self):
        """创建界面"""
        # Excel件选择
        excel_frame = ttk.LabelFrame(self.root, text="Excel文件", padding=5)
        excel_frame.pack(fill=tk.X, padx=5, pady=5)

        self.excel_path = tk.StringVar()
        ttk.Entry(excel_frame, textvariable=self.excel_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(excel_frame, text="选择文件", command=self.select_excel).pack(side=tk.LEFT, padx=5)

        # 账号密码输入
        input_frame = ttk.LabelFrame(self.root, text="登录信息", padding=5)
        input_frame.pack(fill=tk.X, padx=5, pady=5)

        # 账号
        account_frame = ttk.Frame(input_frame)
        account_frame.pack(fill=tk.X, pady=2)
        ttk.Label(account_frame, text="账号：").pack(side=tk.LEFT, padx=5)
        self.account_entry = ttk.Entry(account_frame, width=40)
        self.account_entry.pack(side=tk.LEFT, padx=5)

        # 密码
        password_frame = ttk.Frame(input_frame)
        password_frame.pack(fill=tk.X, pady=2)
        ttk.Label(password_frame, text="密码：").pack(side=tk.LEFT, padx=5)
        self.password_entry = ttk.Entry(password_frame, width=40, show="*")
        self.password_entry.pack(side=tk.LEFT, padx=5)

        # 绑定输入变化事件
        self.account_entry.bind('<FocusOut>', lambda e: self.save_current_config())
        self.password_entry.bind('<FocusOut>', lambda e: self.save_current_config())

        # 行按
        ttk.Button(self.root, text="执行", command=self.execute).pack(pady=10)

        # 日志显示
        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 初始化日志记录器
        self.logger = Logger(self.log_text)


def wait_for_manual_verification(driver, logger, timeout=300):
    """等待用户手动完成验证"""
    logger.log("检测到验证码，请手动完成验证...", "WARNING")
    messagebox.showinfo("提示", "请手动完成验证码，完成后程序将自动继续")

    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # 检查是否已经通过验证（URL不再包含login）
            if "login" not in driver.current_url.lower():
                logger.log("验证通过，继续执行...", "SUCCESS")
                return True
            time.sleep(2)
        except:
            pass

    logger.log("验证码等待时", "ERROR")
    return False


def handle_login(driver, account, password, logger):
    """处理登录流程"""
    try:
        # 等待用名输入框
        logger.log("等待用户名输入框...")
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[2]/div/div[2]/div[2]/input'))
        )
        logger.log("成功找到用户名输入框", "SUCCESS")

        # 输入账号
        logger.log("正在输入账号...")
        type_like_human(username_input, account, logger)
        time.sleep(1)

        # 等待密码输入框
        logger.log("等待密码输入框...")
        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[2]/div/div[2]/input'))
        )
        logger.log("成功找到密码输入框", "SUCCESS")

        # 输入密码
        logger.log("正在输入密码...")
        type_like_human(password_input, password, logger)
        time.sleep(1)

        # 等待登录按钮
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div[2]/div/div[2]/button'))
        )
        logger.log("成功找到登录按钮", "SUCCESS")

        # 点击登录按钮
        logger.log("正在点击录按钮...")
        login_button.click()
        logger.log("已点击登录按钮", "SUCCESS")

        # 等待验证码模块
        time.sleep(2)

        # 检查验证码模块
        logger.log("开始检查验证码模块...", "INFO")
        try:
            # 检查验证码容器
            logger.log("正在查找验证码容器...", "INFO")
            captcha_container = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[2]/div/div[2]'))
            )

            if captcha_container.is_displayed():
                logger.log("验证码容器可见", "INFO")
                current_url = driver.current_url

                # 直接显示提示，不再检查具体的验证码元素
                logger.log("请完成验证码验证", "WARNING")
                messagebox.showinfo(
                    "需要验证",
                    "请手动完成验证\n\n" +
                    "1. 请滑动滑块完成验证\n" +
                    "2. 验证通过后程序将自动继续\n" +
                    "3. 如果失败请重新滑动"
                )

                # 等待验证完成和页面跳转
                logger.log("待验证完成和页面跳转...", "INFO")
                max_wait = 300  # 最长等待5分钟
                wait_start = time.time()

                while time.time() - wait_start < max_wait:
                    # 检查URL是否改变（登录成功）
                    if driver.current_url != current_url:
                        logger.log("验证通过，检测到页面跳转", "SUCCESS")
                        time.sleep(2)  # 等待页面稳定
                        return True
                    time.sleep(1)

                logger.log("验证等待超时", "ERROR")
                return False

        except Exception as e:
            logger.log(f"验证码检测过程出错: {str(e)}", "WARNING")

        # 检查最终状态
        if "login" not in driver.current_url.lower():
            logger.log("登录成功", "SUCCESS")
            return True
        else:
            logger.log("登录失败，仍在登录页面", "ERROR")
            return False

    except Exception as e:
        logger.log(f"登录过程中发生错误: {str(e)}", "ERROR")
        return False


def handle_post_login(driver, excel_path, logger):
    """处理登录后的操作"""
    try:
        # 等待页面完全加载
        logger.log("等待页面元素加载完成...")
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
        logger.log("页面加载完成", "SUCCESS")

        # 给页面一点时间稳定
        time.sleep(3)

        # 记录当前URL用于调试
        current_url = driver.current_url
        logger.log(f"当前页面URL: {current_url}", "INFO")

        # 点击商机沟通菜单
        logger.log("尝试点击商机沟通菜单...", "INFO")
        if not click_business_communication(driver, logger):
            raise Exception("无法进入商机沟通页面")

        # 处理Excel数据
        logger.log("开始处理Excel数据...", "INFO")
        process_excel_data(driver, excel_path, logger)

        logger.log("登录后处理完成", "SUCCESS")
        return True

    except Exception as e:
        logger.log(f"登录后操作失败: {str(e)}", "ERROR")
        return False


def navigate_to_search(driver, logger):
    """导航到搜索页面并准备搜索"""
    try:
        logger.log("等待搜索区域加载...")
        wait = WebDriverWait(driver, 20)

        # 点击搜索区域
        logger.log("在点击搜索区域...")
        search_area = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="widget-27"]/div/form/input'))
        )
        search_area.click()
        time.sleep(1)  # 短暂等待

        # 定位发件人输入框
        logger.log("正在定位发件人输入框...")
        sender_input = wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="search-form-sender"]'))
        )

        # 确保输入框可见且可交互
        if sender_input.is_displayed() and sender_input.is_enabled():
            logger.log("成功找到件输入框", "SUCCESS")
            return sender_input
        else:
            raise Exception("发件人输入框不可用")

    except Exception as e:
        logger.log(f"准备搜索时发生错误: {str(e)}", "ERROR")
        return None


def click_business_communication(driver, logger):
    """点击商机沟通菜单"""
    try:
        logger.log("正在等待商机沟通菜单加载...")
        wait = WebDriverWait(driver, 3)

        # 使用JavaScript移除可能的遮罩层
        try:
            driver.execute_script("""
                var elements = document.querySelectorAll('div[class*="modal"], div[class*="dialog"], div[class*="popup"], div[style*="z-index"]');
                elements.forEach(function(element) {
                    element.remove();
                });
            """)
            logger.log("已尝试移除遮罩层", "INFO")
        except:
            pass

        # 然后尝试点击商机沟通菜单
        menu_element = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="seller-menu-container"]/div/div/div/div[2]/div/ul/div[7]/a'))
        )

        # 使用JavaScript点击，避免被遮挡
        logger.log("正在点击商机沟通菜单...")
        driver.execute_script("arguments[0].click();", menu_element)

        # 等待页面加载完成
        logger.log("等待页面加载完成...")
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        time.sleep(3)  # 额外等待以确保页面完全加载

        # 准备搜索
        sender_input = navigate_to_search(driver, logger)
        if not sender_input:
            raise Exception("无法找到发件人输入框")

        logger.log("商机沟通页面准备完成", "SUCCESS")
        return True

    except Exception as e:
        logger.log(f"点击商机沟通菜单时发生错误: {str(e)}", "ERROR")
        return False


def is_file_locked(filepath):
    """检查文件是否被占用"""
    try:
        with open(filepath, 'ab') as _:
            return False
    except:
        return True


def close_excel_file(file_path, logger):
    """关闭指定的Excel文件"""
    try:
        # 获取Excel应用程序实例
        excel = win32com.client.GetObject(Class="Excel.Application")

        # 检查所有打开的工作簿
        for wb in excel.Workbooks:
            if os.path.abspath(wb.FullName) == os.path.abspath(file_path):
                logger.log(f"找到打开的Excel文件: {file_path}", "INFO")
                # 保存并关闭
                wb.Save()
                wb.Close()
                logger.log("已保存并关闭Excel文件", "SUCCESS")
                return True

        return False
    except:
        return False


def kill_excel_process(logger):
    """强制结束所有Excel进程"""
    try:
        for proc in psutil.process_iter():
            if proc.name().lower() in ['excel.exe', 'xlview.exe']:
                proc.kill()
        logger.log("已闭所有Excel进程", "SUCCESS")
        return True
    except:
        return False

def show_file_locked_dialog(excel_path, logger):
    """显示文件被锁定的提示对话框"""
    try:
        # 创建主窗口
        root = tk.Tk()

        # 设置窗口大小和位置
        window_width = 400
        window_height = 200
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # 将窗口放在左侧中间
        x = 50
        y = (screen_height - window_height) // 2

        root.geometry(f'{window_width}x{window_height}+{x}+{y}')

        # 设置窗口置顶
        root.lift()
        root.attributes('-topmost', True)

        # 隐藏主窗口但保持置顶效果
        root.withdraw()

        # 准备消息内容
        message = ("Excel文件当前被打开，是否保存并关闭？\n\n"
                   "选择'是'：自动保存并关闭Excel\n"
                   "选择'否'：使用临时文件继续\n"
                   "选择'取消'：终止操作")

        # 显示对话框并等待响应
        logger.log("显示文件锁定提示对话框", "INFO")
        response = messagebox.askyesnocancel(
            title="文件被占用",
            message=message,
            icon=messagebox.WARNING
        )

        # 处理响应
        if response is None:  # 取消
            logger.log("用户选择取消操作", "INFO")
            result = "cancel"
        elif response:  # 是
            logger.log("用户选择关闭Excel", "INFO")
            result = "close"
        else:  # 否
            logger.log("用户选择使用临时文件", "INFO")
            result = "temp"

        # 销毁窗口
        root.destroy()
        return result

    except Exception as e:
        logger.log(f"显示对话框时发生错误: {str(e)}", "ERROR")
        return "error"
def process_excel_data(driver, excel_path, logger):
    try:
        # 确保日志窗口在最上层
        if hasattr(logger, 'window'):
            logger.window.lift()
            logger.window.attributes('-topmost', True)

        # 在开始处理前检查文件是否被占用
        if is_file_locked(excel_path):
            logger.log("检测到Excel文件被占用", "WARNING")

            # 显示对话框
            response = show_file_locked_dialog(excel_path, logger)

            if response == "cancel":
                logger.log("操作已取消", "INFO")
                return False
            elif response == "close":
                logger.log("尝试关闭Excel文件...", "INFO")
                if close_excel_file(excel_path, logger):
                    logger.log("成功关闭Excel文件", "SUCCESS")
                    time.sleep(1)  # 等待文件释放
                else:
                    logger.log("无法正常关闭Excel，尝试强制关闭...", "WARNING")
                    if kill_excel_process(logger):
                        logger.log("已强制关闭Excel", "SUCCESS")
                        time.sleep(1)  # 等待文件释放
                    else:
                        logger.log("无法关闭Excel，将使用临时文件", "WARNING")
                        response = "temp"

            if response == "temp":
                logger.log("将使用临时文件机制继续", "INFO")
                # 更新文件路径为临时文件
                temp_path = excel_path.rsplit('.', 1)[0] + '_temp.' + excel_path.rsplit('.', 1)[1]
                excel_path = temp_path
                logger.log(f"临时文件路径: {excel_path}", "INFO")

            elif response == "error":
                logger.log("对话框显示失败，操作取消", "ERROR")
                return False

        # 继续处理Excel数据
        # 读取Excel文件
        logger.log("正在读取Excel文件...")
        df = pd.read_excel(excel_path, engine='xlrd')  # 使用 xlrd 引擎读取 .xls
        logger.log("成功读取Excel文件", "SUCCESS")

        # 确保P列存在
        if df.shape[1] < 16:
            # 添加缺失的列
            for i in range(df.shape[1], 16):
                col_name = chr(ord('A') + i)
                df[col_name] = ''
            logger.log("已添加缺失的列", "INFO")

        # 定位发件人输入框
        sender_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="search-form-sender"]'))
        )

        # 从D7开始处理每一行
        start_row = 6  # Excel中的第7行
        processed_count = 0  # 记录处理的行数
        skipped_count = 0  # 记录跳过的行数

        try:
            while start_row < len(df):
                # 检查P列是否已有数据
                p_column_value = str(df.iloc[start_row, 15]).strip()  # P列的索引是15
                if pd.notna(p_column_value) and p_column_value != '' and p_column_value != 'nan':
                    logger.log(f"跳过第{start_row + 1}行：P列已有数据 '{p_column_value}'", "INFO")
                    skipped_count += 1
                    start_row += 1
                    continue

                name = df.iloc[start_row, 3]  # D列的索引是3
                if pd.isna(name):  # 如果是空值就跳过
                    logger.log(f"跳过第{start_row + 1}行：D列为空", "INFO")
                    skipped_count += 1
                    start_row += 1
                    continue

                logger.log(f"正在处理第{start_row + 1}行: {name}")

                # 直接清空并填入内容
                sender_input.clear()
                sender_input.send_keys(Keys.CONTROL + "a", Keys.DELETE)
                sender_input.send_keys(str(name))

                # 点击搜索按钮
                search_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="widget-41"]/form/div[4]/button[1]'))
                )
                search_button.click()

                # 等待搜索结果加载
                logger.log("等待搜索结果加载...")
                time.sleep(1.5)

                try:
                    # 点击链接打开新窗口
                    link = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="widget-29"]/div[1]'))
                    )
                    link.click()

                    # 切换到新窗口
                    new_window = driver.window_handles[-1]
                    driver.switch_to.window(new_window)

                    # 等待新页面中的元素加载
                    content_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '//*[@id="app-crm"]/div/div/div[2]/div/div[2]/div/div[2]/div[2]'))
                    )

                    # 获取内容
                    content = content_element.text
                    logger.log(f"获取到内容: {content}")

                    # 关闭新窗口
                    driver.close()

                    # 切回主窗口
                    driver.switch_to.window(driver.window_handles[0])

                    # 写入内容并保存
                    df.iloc[start_row, 15] = content
                    if not save_excel_data(df, excel_path, logger):
                        logger.log("保存失败，创建备份...", "WARNING")
                        # 创建备份
                        backup_path = excel_path.rsplit('.', 1)[0] + '_backup.' + excel_path.rsplit('.', 1)[1]
                        save_excel_data(df, backup_path, logger)

                    logger.log(f"已将内容'{content}'保存到第{start_row + 1}行P列", "SUCCESS")
                    processed_count += 1

                except Exception as e:
                    logger.log(f"处理第{start_row + 1}行时发生错误: {str(e)}", "ERROR")

                start_row += 1
                time.sleep(1)

        except KeyboardInterrupt:
            logger.log("检测到手动中断，正在保存当前进度...", "WARNING")
            save_excel_data(df, excel_path, logger)
            raise

        # 最后的统计
        logger.log("================================================================", "SUCCESS")
        logger.log("                      处理完成！", "SUCCESS")
        logger.log("----------------------------------------------------------------", "SUCCESS")
        logger.log(f"总共处理: {processed_count}行", "SUCCESS")
        logger.log(f"总共跳过: {skipped_count}行", "SUCCESS")
        logger.log(f"总行数: {len(df)}行", "SUCCESS")
        logger.log("================================================================", "SUCCESS")

    except Exception as e:
        logger.log(f"处理Excel数据时发生错误: {str(e)}", "ERROR")
        raise


def execute_action(excel_path, account, password, logger):
    """执行主要操作"""
    try:
        if not all([excel_path, account, password]):
            logger.log("错误：请填写所有必要信息！", "ERROR")
            return

        logger.log("正在检查配置...")
        if not check_paths(logger):
            return

        # 检查网络连接
        logger.log("正在检查网络连接...")
        if not check_internet_connection(logger):
            messagebox.showerror("错误", "网络连接失败，请检查网络设置！")
            return

        # 获取Chrome配置
        logger.log("正在配置Chrome浏览器...")
        chrome_options = get_chrome_config()

        # 获取ChromeDriver路径
        driver_path = get_driver_path(logger)
        if not driver_path:
            return

        try:
            logger.log("正在启动Chrome浏览器...")
            service = Service(driver_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)

            # 设置超时时间
            driver.set_page_load_timeout(60)
            driver.implicitly_wait(20)

            # 直接访问登录页面
            logger.log("正在打开登录页...")
            target_url = "https://i.alibaba.com/index.htm"
            driver.get(target_url)

            # 处理登录
            if not handle_login(driver, account, password, logger):
                raise Exception("登录失败")

            # 登录成功后，等待一下再处理后续操作
            time.sleep(3)

            # 处理登录后的操作
            if not handle_post_login(driver, excel_path, logger):
                logger.log("继续执行后续操作...", "INFO")
                # 即使post_login失败也继续执行
                pass

            # 继续执行其他操作...
            logger.log("所有操作完成", "SUCCESS")

        except Exception as e:
            error_msg = f"浏览器操作发生错误：{str(e)}"
            logger.log(error_msg, "ERROR")
            messagebox.showerror("错误", error_msg)

        finally:
            try:
                if 'driver' in locals():
                    driver.quit()
            except:
                pass

    except Exception as e:
        error_msg = f"执行过程中发生错误：{str(e)}"
        logger.log(error_msg, "ERROR")
        messagebox.showerror("错误", error_msg)


def check_paths(logger):
    """检查必要的径是否存在"""
    chrome_path = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'

    if not os.path.exists(chrome_path):
        error_msg = f"Chrome浏览器路径不存在: {chrome_path}\n请检查Chrome是否正确安装！"
        logger.log(error_msg, "ERROR")
        messagebox.showerror("错误", error_msg)
        return False

    logger.log("所有路径检查通过", "SUCCESS")
    return True
def save_excel_data(df, excel_path, logger):
    """保存Excel数据，处理文件被占用的情况"""
    try:
        # 检查文件类型
        is_xls = excel_path.lower().endswith('.xls')

        if is_xls:
            # 对于 .xls 文件，使用 xlwt
            import xlwt
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Sheet1')

            # 写入列名
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, str(col_name))

            # 写入数据
            for row_idx in range(len(df)):
                for col_idx in range(len(df.columns)):
                    value = df.iloc[row_idx, col_idx]
                    # 处理空值
                    if pd.isna(value):
                        value = ''
                    ws.write(row_idx + 1, col_idx, str(value))

            # 保存文件
            try:
                wb.save(excel_path)
                logger.log(f"成功保存数据到: {excel_path}", "SUCCESS")
                return True
            except Exception as save_error:
                # 如果保存失败，尝保存为临时文件
                temp_path = excel_path.rsplit('.', 1)[0] + '_temp.xls'
                wb.save(temp_path)
                logger.log(f"已保存到临时文件: {temp_path}", "WARNING")
                return True

        else:
            # 对于其他格式（如 .xlsx），使用 pandas
            df.to_excel(excel_path, index=False, engine='openpyxl')
            logger.log(f"成功保存数据到: {excel_path}", "SUCCESS")
            return True

    except Exception as e:
        logger.log(f"保存Excel时发生错误: {str(e)}", "ERROR")
        # 尝试创建备份
        try:
            backup_path = excel_path.rsplit('.', 1)[0] + '_backup.xls'
            if is_xls:
                # 使用 xlwt 保存备份
                wb = xlwt.Workbook()
                ws = wb.add_sheet('Sheet1')
                for col_idx, col_name in enumerate(df.columns):
                    ws.write(0, col_idx, str(col_name))
                for row_idx in range(len(df)):
                    for col_idx in range(len(df.columns)):
                        value = df.iloc[row_idx, col_idx]
                        if pd.isna(value):
                            value = ''
                        ws.write(row_idx + 1, col_idx, str(value))
                wb.save(backup_path)
            else:
                df.to_excel(backup_path, index=False)
            logger.log(f"已创建备份文件: {backup_path}", "WARNING")
            return True
        except:
            logger.log("创建备份文件失败", "ERROR")
            return False

def type_like_human(element, text, logger):
    """快速输入文本"""
    try:
        # 清空输入框
        element.clear()
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)

        # 直接输入文本
        element.send_keys(text)
        return True
    except Exception as e:
        logger.log(f"输入文本时发生错误: {str(e)}", "ERROR")
        return False


def init_driver(logger):
    """初始化浏览器"""
    try:
        # 获取Chrome配置
        options = get_chrome_config()

        # 创建WebDriver实例
        driver = webdriver.Chrome(options=options)

        # 设置窗口位置
        screen_width = driver.execute_script('return window.screen.width')
        x = screen_width - 1020
        driver.set_window_position(x, 20)

        # 设置超时
        driver.set_page_load_timeout(30)
        driver.implicitly_wait(10)

        logger.log("浏览器初始化成功", "SUCCESS")
        return driver

    except Exception as e:
        logger.log(f"浏览器初始化失败: {str(e)}", "ERROR")
        raise


def main():
    root = tk.Tk()
    app = Application(root)
    root.mainloop()


if __name__ == "__main__":
    main()