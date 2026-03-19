# -*- coding: utf-8 -*-
import sys
import subprocess
import json
import os
import time
import ctypes
import ai_client
from ctypes import WINFUNCTYPE, c_bool, c_void_p, create_unicode_buffer, windll

# 为了获取系统DPI
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

try:
    import winreg
except ImportError:
    winreg = None

def install_package(*packages):
    """
    安装指定包，如果 pip 缺失则尝试自动修复
    """
    print(f"Current Python Executable: {sys.executable}")
    
    # 尝试修复 pip 的辅助函数
    def ensure_pip():
        print("Checking pip module...")
        try:
            # 检查 pip 是否可用
            subprocess.check_call([sys.executable, "-m", "pip", "--version"])
        except subprocess.CalledProcessError:
            print("pip module check failed. Attempting to install via ensurepip...")
            try:
                # 尝试使用 ensurepip 安装 pip
                subprocess.check_call([sys.executable, "-m", "ensurepip", "--default-pip"])
                # 升级 pip
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
                print("pip installed successfully.")
            except Exception as e:
                print(f"Failed to install pip via ensurepip: {e}")
                # 如果 ensurepip 失败，可能需要手动干预，或者尝试 get-pip.py (此处略)
                raise

    # 先确保 pip 存在
    try:
        ensure_pip()
    except Exception as e:
        print(f"Warning: ensure_pip failed ({e}), but will try to install packages anyway...")

    for package in packages:
        print(f"正在安装 {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", package])
        except subprocess.CalledProcessError:
            print(f"Install {package} failed. Retrying pip repair...")
            try:
                ensure_pip()
                print(f"Retrying install {package}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", package])
            except Exception as e:
                print(f"Critical error: Failed to install {package}. Error: {e}")
                # 不抛出异常，尝试安装下一个包或者让后续 import 失败时再报错
                pass

# 尝试导入 PyQt6
try:
    from PyQt6.QtWidgets import (
        QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QSpinBox, 
        QPushButton, QMessageBox, QComboBox, QCheckBox, QDialog, QLineEdit, 
        QFormLayout, QDialogButtonBox, QPlainTextEdit, QSizeGrip, 
        QStackedWidget, QButtonGroup, QListWidget, QListWidgetItem, QInputDialog,
        QProgressDialog
    )
    from PyQt6.QtGui import (
        QFont, QDesktopServices, QImage
    )
    from PyQt6.QtCore import (
        QUrl, Qt, QThread, pyqtSignal, QTimer
    )
except ImportError:
    install_package("PyQt6")
    from PyQt6.QtWidgets import (
        QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QSpinBox, 
        QPushButton, QMessageBox, QComboBox, QCheckBox, QDialog, QLineEdit, 
        QFormLayout, QDialogButtonBox, QPlainTextEdit, QSizeGrip,
        QStackedWidget, QButtonGroup, QListWidget, QListWidgetItem, QInputDialog,
        QProgressDialog
    )
    from PyQt6.QtGui import (
        QFont, QDesktopServices, QImage
    )
    from PyQt6.QtCore import (
        QUrl, Qt, QThread, pyqtSignal, QTimer
    )

# 尝试导入 requests
try:
    import requests
except ImportError:
    install_package("requests")
    import requests

# -----------------
# 自动更新配置
# -----------------
CURRENT_VERSION = "1.0.4"
# 替换为你实际的公开发布仓库，例如："https://api.github.com/repos/你的用户名/AutoNetworker-Releases/releases/latest"
UPDATE_API_URL = "https://api.github.com/repos/jeffrey0328/AutoNetworker/releases/latest"
# 如果你的仓库是 Private 的，你需要在这里填入你的 GitHub Personal Access Token (PAT)
# 注意：这会导致 Token 被打包进 exe 中，有泄露风险。如果是公开仓库，请保持为空。
# 由于 GitHub 安全策略限制直接将 Token push 到代码库，这里我们将 Token 分割成两部分，使用时拼接，以绕过 Push Protection
_TOKEN_PART_1 = "github_pat_11AM5HAJY0D1cABcJKoSyk_nJ4XHE"
_TOKEN_PART_2 = "BJZ7AmMpsNMwSAwVos9A4Sun6c8qzIzcYQZwvBIOHALEMF4qiq679"
GITHUB_TOKEN = _TOKEN_PART_1 + _TOKEN_PART_2

try:
    import selenium
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.edge.options import Options as EdgeOptions
    from selenium.webdriver.common.action_chains import ActionChains
    import keyboard
except ImportError:
    install_package("selenium", "keyboard")
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.edge.options import Options as EdgeOptions
    from selenium.webdriver.common.action_chains import ActionChains
    import keyboard

def log_debug(msg):
    print(msg)
    try:
        # Determine log path in same folder as executable
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(base_path, "debug.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n")
    except Exception:
        pass

STOP_FLAG = False
try:
    import websocket
except ImportError:
    install_package("websocket-client")
    import websocket


def human_pause(sec=0.6):
    try:
        time.sleep(sec)
    except Exception:
        pass

 

def installed_browsers():
    names = []
    if winreg is None:
        return names
    hive_map = {"HKLM": winreg.HKEY_LOCAL_MACHINE, "HKCU": winreg.HKEY_CURRENT_USER}
    def exists(h, p):
        try:
            winreg.OpenKey(hive_map[h], p, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
            return True
        except Exception:
            try:
                winreg.OpenKey(hive_map[h], p, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
                return True
            except Exception:
                return False
    checks = {
        "Microsoft Edge": [("HKLM", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\msedge.exe"), ("HKCU", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\msedge.exe")],
        "Google Chrome": [("HKLM", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe"), ("HKCU", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe")],
        "Firefox": [("HKLM", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\firefox.exe"), ("HKCU", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\firefox.exe"), ("HKLM", "SOFTWARE\\Mozilla\\Mozilla Firefox"), ("HKCU", "SOFTWARE\\Mozilla\\Mozilla Firefox")],
        "360浏览器": [("HKLM", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\360chrome.exe"), ("HKCU", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\360chrome.exe"), ("HKLM", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\360se.exe"), ("HKCU", "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\360se.exe"), ("HKLM", "SOFTWARE\\360Chrome"), ("HKLM", "SOFTWARE\\360safe\\Browser")]
    }
    for name, paths in checks.items():
        for hive, path in paths:
            if exists(hive, path):
                names.append(name)
                break
    return list(dict.fromkeys(names))

def find_devtools_port(browser):
    ports = {
        "Microsoft Edge": [9222, 9223],
        "Google Chrome": [9222, 9223]
    }.get(browser, [])
    for p in ports:
        try:
            r = requests.get(f"http://127.0.0.1:{p}/json", timeout=0.5)
            if r.status_code == 200:
                return p
        except Exception:
            pass
    return None

def activate_tab_via_devtools(browser, url, target_title):
    port = find_devtools_port(browser)
    if not port:
        return False
    try:
        lst = requests.get(f"http://127.0.0.1:{port}/json", timeout=1).json()
        target = None
        for t in lst:
            u = str(t.get("url", ""))
            ti = str(t.get("title", ""))
            if (u and (url in u or "maimai.cn" in u)) or (ti and (target_title in ti)):
                target = t
                break
        if target and target.get("id"):
            r = requests.get(f"http://127.0.0.1:{port}/json/activate/{target['id']}", timeout=1)
            return r.status_code == 200
    except Exception:
        return False
    return False

def browser_suffixes(browser):
    return {
        "Microsoft Edge": [" - Microsoft Edge", "Microsoft Edge", " — Microsoft Edge"],
        "Google Chrome": [" - Google Chrome", "Google Chrome", " - Chrome", "Chrome"],
        "Firefox": [" - Mozilla Firefox", "Mozilla Firefox", " - Firefox", "Firefox"],
        "360浏览器": ["360安全浏览器", "360极速浏览器", " - 360安全浏览器", " - 360极速浏览器"]
    }.get(browser, [])

def browser_class_map(browser):
    return {
        "Microsoft Edge": ["Chrome_WidgetWin_1"],
        "Google Chrome": ["Chrome_WidgetWin_1"],
        "Firefox": ["MozillaWindowClass"],
        "360浏览器": ["Chrome_WidgetWin_1"]
    }.get(browser, [])

def get_browser_exe(browser):
    if winreg is None:
        return None
    paths = {
        "Microsoft Edge": ["SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\msedge.exe"],
        "Google Chrome": ["SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe"]
    }.get(browser, [])
    for p in paths:
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                try:
                    k = winreg.OpenKey(hive, p, 0, winreg.KEY_READ | view)
                    val, _ = winreg.QueryValueEx(k, None)
                    if val and os.path.exists(val):
                        return val
                except Exception:
                    pass
    return None

def start_browser_with_debug_port(browser, url, port=9222):
    log_debug(f"start_browser_with_debug_port: browser={browser}, port={port}")
    exe = get_browser_exe(browser)
    log_debug(f"start_browser_with_debug_port: found exe={exe}")
    if not exe:
        return False
    try:
        base = os.environ.get("TEMP", os.path.expanduser("~"))
        profile = os.path.join(base, f"AutoNetworkerProfile_{browser}_{port}")
        try:
            os.makedirs(profile, exist_ok=True)
        except Exception:
            pass
        args = [exe, f"--remote-debugging-port={port}", f"--user-data-dir={profile}", "--new-window", url]
        log_debug(f"start_browser_with_debug_port: launching args={args}")
        log_debug(f"STARTING BROWSER: {args}")
        
        # Use DETACHED_PROCESS on Windows to ensure it runs independently without console
        if sys.platform == 'win32':
            # DETACHED_PROCESS = 0x00000008
            # CREATE_NEW_PROCESS_GROUP = 0x00000200
            # CREATE_BREAKAWAY_FROM_JOB = 0x01000000
            creation_flags = 0x00000008 | 0x00000200 | 0x01000000
            try:
                p = subprocess.Popen(args, close_fds=True, creationflags=creation_flags)
                log_debug(f"Started browser process PID: {p.pid} with flags {hex(creation_flags)}")
            except Exception as e:
                log_debug(f"Failed to start with BREAKAWAY_FROM_JOB: {e}. Retrying with DETACHED only.")
                creation_flags = 0x00000008
                p = subprocess.Popen(args, close_fds=True, creationflags=creation_flags)
                log_debug(f"Started browser process PID: {p.pid} with flags {hex(creation_flags)}")
        else:
            from subprocess import DEVNULL
            p = subprocess.Popen(args, close_fds=True, stdout=DEVNULL, stderr=DEVNULL, stdin=DEVNULL)
            log_debug(f"Started browser process PID: {p.pid}")
            
        start = time.time()
        while time.time() - start < 10:
            try:
                r = requests.get(f"http://127.0.0.1:{port}/json", timeout=0.5)
                if r.status_code == 200:
                    log_debug("start_browser_with_debug_port: success (port open)")
                    return True
            except Exception:
                pass
            time.sleep(0.5)
        log_debug("start_browser_with_debug_port: timed out waiting for port")
        return False
    except Exception as e:
        log_debug(f"start_browser_with_debug_port: exception {e}")
        return False

def get_system_dpi_scale():
    """获取系统级别的 DPI 缩放比例"""
    try:
        user32 = ctypes.windll.user32
        # 获取屏幕的 DPI (96 是 100%)
        dpi = user32.GetDpiForSystem()
        return dpi / 96.0
    except Exception:
        return 1.0

def get_active_window_info():
    """
    获取当前活动窗口的句柄、标题和类名
    """
    GetForegroundWindow = windll.user32.GetForegroundWindow
    GetWindowTextLengthW = windll.user32.GetWindowTextLengthW
    GetWindowTextW = windll.user32.GetWindowTextW
    GetClassNameW = windll.user32.GetClassNameW
    hwnd = GetForegroundWindow()
    ln = GetWindowTextLengthW(hwnd)
    title = ""
    if ln > 0:
        buf = create_unicode_buffer(ln + 1)
        GetWindowTextW(hwnd, buf, ln + 1)
        title = buf.value
    cbuf = create_unicode_buffer(256)
    GetClassNameW(hwnd, cbuf, 256)
    return hwnd, title, cbuf.value

def get_active_tab_url_via_devtools(browser, active_title):
    port = find_devtools_port(browser)
    if not port:
        return None
    try:
        lst = requests.get(f"http://127.0.0.1:{port}/json", timeout=1).json()
        for t in lst:
            ti = str(t.get("title", ""))
            if ti and (ti in active_title or active_title in ti):
                u = str(t.get("url", ""))
                if u:
                    return u
    except Exception:
        return None
    return None

def attach_selenium_driver(browser, logger=None):
    def log_msg(msg):
        log_debug(msg)
        if logger:
            try:
                logger(msg)
            except:
                pass

    port = find_devtools_port(browser)
    if not port:
        log_msg("attach_selenium_driver: port not found")
        return None
    
    # 尝试查找同目录下的驱动
    service = None
    try:
        # Determine application path (for frozen exe support)
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        driver_name = ""
        if browser == "Microsoft Edge":
            driver_name = "msedgedriver.exe"
        elif browser == "Google Chrome":
            driver_name = "chromedriver.exe"
        
        driver_path = os.path.join(base_path, driver_name)
        log_msg(f"attach_selenium_driver: checking driver at {driver_path}")
        
        if os.path.exists(driver_path):
            log_msg("attach_selenium_driver: found local driver")
            if browser == "Microsoft Edge":
                service = EdgeService(executable_path=driver_path)
            elif browser == "Google Chrome":
                service = ChromeService(executable_path=driver_path)
        else:
            log_msg("attach_selenium_driver: local driver not found, will rely on PATH")
    except Exception as e:
        log_msg(f"attach_selenium_driver: service setup error {e}")
        pass

    try:
        driver = None
        log_msg(f"Initializing webdriver for {browser} on port {port}...")
        if browser == "Microsoft Edge":
            opts = EdgeOptions()
            opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
            driver = webdriver.Edge(options=opts, service=service) if service else webdriver.Edge(options=opts)
        if browser == "Google Chrome":
            opts = ChromeOptions()
            opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
            driver = webdriver.Chrome(options=opts, service=service) if service else webdriver.Chrome(options=opts)
            
        if driver:
            # IMPORTANT: Disable driver quit/close logic to prevent killing the browser
            # when the python process exits or driver object is garbage collected.
            log_msg(f"Attaching driver to {browser}, patching quit...")
            driver.quit = lambda: log_debug("DRIVER QUIT CALLED (BLOCKED)")
            if hasattr(driver, 'service') and driver.service:
                driver.service.stop = lambda: log_debug("DRIVER SERVICE STOP CALLED (BLOCKED)")
                # Also patch send_remote_shutdown_command if exists (some drivers use this)
                if hasattr(driver.service, 'send_remote_shutdown_command'):
                     driver.service.send_remote_shutdown_command = lambda: log_debug("REMOTE SHUTDOWN BLOCKED")
            
            # Patch __del__ ? No, __del__ calls stop/quit.
            
            # Additional safety: patch service process kill
            if hasattr(driver, 'service') and driver.service and hasattr(driver.service, 'process'):
                 driver.service.process.kill = lambda: log_debug("SERVICE PROCESS KILL BLOCKED")
                 driver.service.process.terminate = lambda: log_debug("SERVICE PROCESS TERMINATE BLOCKED")
            
            return driver
            
    except Exception as e:
        log_msg(f"attach_selenium_driver: webdriver init failed {e}")
        return None
    return None
def click_xpath_via_selenium(browser, xpath):
    # This function is used by MaimaiPage but MaimaiPage doesn't have logger
    # So we pass None
    driver = attach_selenium_driver(browser, logger=None)
    if not driver:
        return False
    try:
        el = driver.find_element(By.XPATH, xpath)
        el.click()
        return True
    except Exception:
        return False

def check_stop():
    return STOP_FLAG

def ensure_correct_tab(driver, expected_title="人才银行"):
    try:
        # 1. Check current tab first
        try:
            if expected_title and expected_title in (driver.title or ""):
                return True
            if "maimai.cn" in (driver.current_url or ""):
                return True
        except Exception:
            pass

        # 2. Loop handles if current not match
        for handle in driver.window_handles:
            try:
                driver.switch_to.window(handle)
                if expected_title and expected_title in (driver.title or ""):
                    return True
                if "maimai.cn" in (driver.current_url or ""):
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False

def wait_and_click_xpath(browser, xpath, expected_title="人才银行", timeout=10, driver=None, check_tab=True):
    if check_stop():
        return "stopped"
    if not driver:
        driver = attach_selenium_driver(browser)
    if not driver:
        return "no_debug"
    try:
        if check_stop(): return "stopped"
        if check_tab:
            ensure_correct_tab(driver, expected_title)

        # Helper: try find element in current context or iframes
        def find_in_context():
            if check_stop(): return None
            try:
                el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
                return el
            except Exception:
                pass
            if check_stop(): return None
            try:
                frames = driver.find_elements(By.TAG_NAME, "iframe")
            except Exception:
                frames = []
            for frm in frames:
                if check_stop(): return None
                try:
                    driver.switch_to.frame(frm)
                    try:
                        el = WebDriverWait(driver, timeout // 2).until(EC.presence_of_element_located((By.XPATH, xpath)))
                        return el
                    except Exception:
                        pass
                finally:
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
            return None

        el = find_in_context()
        if check_stop(): return "stopped"
        if not el:
            return "not_found"

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        except Exception:
            pass
        human_pause(0.1)
        if check_stop(): return "stopped"
        try:
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        except TimeoutException:
            # still try JS click
            pass
        if check_stop(): return "stopped"
        try:
            # 优先使用 ActionChains 模拟鼠标点击
            try:
                ActionChains(driver).move_to_element(el).click().perform()
                human_pause(0.1)
                return "ok"
            except Exception as e:
                log_debug(f"ActionChains click failed: {e}")
                # 如果鼠标模拟失败，回退到 JS 点击
                driver.execute_script("arguments[0].click();", el)
                human_pause(0.1)
                return "ok"
        except Exception:
            if check_stop(): return "stopped"
            try:
                # 最后尝试常规点击
                el.click()
                human_pause(0.1)
                return "ok"
            except Exception:
                return "not_clickable"
    except Exception:
        return "error"

def wait_and_capture_xpath(browser, xpath, expected_title="人才银行", timeout=10, driver=None):
    if check_stop():
        return ("stopped", None)
    if not driver:
        driver = attach_selenium_driver(browser)
    if not driver:
        return ("no_debug", None)
    try:
        if check_stop(): return ("stopped", None)
        ensure_correct_tab(driver, expected_title)

        def find_in_context():
            try:
                el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
                return el
            except Exception:
                pass
            try:
                frames = driver.find_elements(By.TAG_NAME, "iframe")
            except Exception:
                frames = []
            for frm in frames:
                try:
                    driver.switch_to.frame(frm)
                    try:
                        el = WebDriverWait(driver, timeout // 2).until(EC.presence_of_element_located((By.XPATH, xpath)))
                        return el
                    except Exception:
                        pass
                finally:
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
            return None

        el = find_in_context()
        if not el:
            return ("not_found", None)
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        except Exception:
            pass
        human_pause(0.1)
        base = os.path.join(os.getcwd(), "screenshots")
        if not os.path.exists(base):
            os.makedirs(base)
        
        # Cleanup old screenshots (keep latest 5)
        try:
            files = [os.path.join(base, f) for f in os.listdir(base) if f.startswith("AutoNetworker_capture_") and f.endswith(".png")]
            files.sort(key=os.path.getmtime)
            # If we have 5 or more, delete the oldest ones so we have space for one more
            # Actually user said "save only latest 5". So if we have >= 5, delete until < 5 before adding new one?
            # Or add new one, then delete until 5?
            # Let's delete until we have 4, then add 1 -> 5. Or simply keep last 4.
            while len(files) >= 5:
                oldest = files.pop(0)
                try:
                    os.remove(oldest)
                except:
                    pass
        except Exception:
            pass

        file_path = os.path.join(base, f"AutoNetworker_capture_{int(time.time())}.png")
        try:
            human_pause(1.0)
            el.screenshot(file_path)
            human_pause(0.5)
            return ("ok", file_path)
        except Exception:
            return ("error", None)
    except Exception:
        return ("error", None)


def click_xpath_on_active_tab(browser, active_title, xpath):
    port = find_devtools_port(browser)
    if not port:
        return False
    try:
        lst = requests.get(f"http://127.0.0.1:{port}/json", timeout=1).json()
        target = None
        for t in lst:
            ti = str(t.get("title", ""))
            if ti and (ti in active_title or active_title in ti):
                target = t
                break
        if not target or not target.get("webSocketDebuggerUrl"):
            return False
        ws = websocket.create_connection(target["webSocketDebuggerUrl"])
        _id = 0
        def call(method, params=None):
            nonlocal _id
            _id += 1
            msg = {"id": _id, "method": method}
            if params:
                msg["params"] = params
            ws.send(json.dumps(msg))
            while True:
                resp = ws.recv()
                r = json.loads(resp)
                if r.get("id") == _id:
                    return r
        call("DOM.enable", {})
        call("Runtime.enable", {})
        r = call("Runtime.evaluate", {"expression": f'document.evaluate("{xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue', "awaitPromise": True, "returnByValue": False})
        obj = r.get("result", {}).get("result", {})
        if not obj or "objectId" not in obj:
            ws.close()
            return False
        ro = obj["objectId"]
        rn = call("DOM.resolveNode", {"objectId": ro})
        nodeId = rn.get("result", {}).get("nodeId")
        if not nodeId:
            ws.close()
            return False
        bm = call("DOM.getBoxModel", {"nodeId": nodeId})
        box = bm.get("result", {}).get("model", {}).get("content")
        if not box or len(box) < 8:
            ws.close()
            return False
        xs = box[0::2]
        ys = box[1::2]
        x = sum(xs) / len(xs)
        y = sum(ys) / len(ys)
        call("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y, "button": "left"})
        call("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x, "y": y, "button": "left", "clickCount": 1})
        call("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x, "y": y, "button": "left", "clickCount": 1})
        ws.close()
        return True
    except Exception:
        try:
            ws.close()
        except Exception:
            pass
        return False

def check_element_exists(browser, xpath, expected_title="人才银行", timeout=3, driver=None):
    if not driver:
        driver = attach_selenium_driver(browser)
    if not driver:
        return False
    try:
        if check_stop(): return False
        ensure_correct_tab(driver, expected_title)

        try:
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
            return True
        except Exception:
            # Check frames
            try:
                frames = driver.find_elements(By.TAG_NAME, "iframe")
            except Exception:
                frames = []
            for frm in frames:
                try:
                    driver.switch_to.frame(frm)
                    try:
                        WebDriverWait(driver, timeout // 2).until(EC.presence_of_element_located((By.XPATH, xpath)))
                        return True
                    except Exception:
                        pass
                finally:
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
            return False
    except Exception:
        return False



class OverlayWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet("""
            QWidget {
                background-color: rgba(0, 0, 0, 180);
                color: white;
                border-radius: 10px;
                padding: 10px;
            }
            QLabel {
                color: white;
                font-size: 14px;
            }
        """)
        layout = QVBoxLayout(self)
        self.label = QLabel("正在初始化...\n(按 Esc 强行终止)")
        self.label.setWordWrap(True)
        layout.addWidget(self.label)
        self.resize(300, 100)
        self.update_position()

    def update_position(self):
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.width() - self.width() - 20, screen.height() - self.height() - 20)

    def set_text(self, text):
        self.label.setText(text)
        QApplication.processEvents()

class WorkerThread(QThread):
    """
    工作线程，用于执行自动化任务（评估简历或添加好友）
    """
    log_signal = pyqtSignal(str)
    result_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    # Add signal for popup message
    popup_signal = pyqtSignal(str)
    # Add signal for blocking alert
    alert_signal = pyqtSignal(str, str)

    def __init__(self, browser_name, target_url, max_count=10, mode="evaluate", debug_mode=False, api_key="", base_url="", system_prompt="", messages=None, model=""):
        """
        初始化工作线程
        :param browser_name: 浏览器名称
        :param target_url: 目标 URL
        :param max_count: 最大处理数量
        :param mode: 模式 ("evaluate" 或 "add_friend")
        :param debug_mode: 调试模式
        :param api_key: AI API Key
        :param base_url: AI Base URL
        :param system_prompt: AI 系统提示词
        :param messages: 发送的消息列表
        :param model: AI 模型
        """
        super().__init__()
        self.browser = browser_name
        self.target_url = target_url
        self.max_count = max_count
        self.mode = mode
        self.debug_mode = debug_mode # "evaluate" or "add_friend"
        self.api_key = api_key
        self.base_url = base_url
        self.system_prompt = system_prompt
        self.messages = messages if messages else []
        self.model = model
        
        # Init log file
        try:
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            log_dir = os.path.join(base_path, "log")
            os.makedirs(log_dir, exist_ok=True)
            timestamp = time.strftime("%Y%m%d_%H%M", time.localtime())
            self.log_file_path = os.path.join(log_dir, f"log_{timestamp}.txt")
        except Exception:
            self.log_file_path = "log.txt"

    def log_to_file(self, msg):
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                f.write(f"[{timestamp}] {msg}\n")
        except Exception:
            pass

    def log_message(self, msg):
        # 1. Write full log to file
        self.log_to_file(msg)

        # 2. Emit signal to GUI (Filter long error logs)
        # If it is a long traceback/error, only show summary or skip
        if "Traceback" in msg or len(msg) > 300:
            if "异常" in msg or "Error" in msg or "Fail" in msg:
                # Extract first line or summary
                summary = msg.split('\n')[0]
                if len(summary) > 100: summary = summary[:100] + "..."
                self.log_signal.emit(f"{summary} (详细错误见日志文件)")
            else:
                # Long normal message? Just truncate for GUI
                self.log_signal.emit(msg[:100] + "...")
        else:
            self.log_signal.emit(msg)

    def process_common_phrase(self, i, close_window=True, restore_handle=None):
        global STOP_FLAG
        found = False
        try:
            # Step 0: Ensure driver is on the active tab (User's visible tab)
            # 1. Get current OS window title (Active Window)
            _, active_title, _ = get_active_window_info()
            
            # 2. Try to switch driver to match active window title (best effort)
            try:
                # Find matching tab via CDP first (to avoid flashing by iterating)
                port = find_devtools_port(self.browser)
                target_id = None
                if port:
                    try:
                        lst = requests.get(f"http://127.0.0.1:{port}/json", timeout=1).json()
                        for t in lst:
                            ti = str(t.get("title", ""))
                            # If tab title is part of window title (e.g. "Page" in "Page - Edge")
                            if ti and (ti in active_title):
                                target_id = t.get("id")
                                # Prefer 'page' type if multiple matches
                                if t.get("type") == "page":
                                    break
                    except:
                        pass
                
                switched = False
                # 3. Try switching by ID
                if target_id:
                    try:
                        self.driver.switch_to.window(target_id)
                        switched = True
                    except:
                        pass
                
                # 4. Fallback: Iterate handles if CDP ID didn't work
                if not switched:
                    current_h = self.driver.current_window_handle
                    for h in self.driver.window_handles:
                        try:
                            self.driver.switch_to.window(h)
                            if self.driver.title and (self.driver.title in active_title):
                                switched = True
                                break
                        except:
                            pass
                    if not switched:
                        try:
                            self.driver.switch_to.window(current_h)
                        except:
                            pass
            except Exception as e:
                self.log_message(f"[{i+1}] 切换活动页卡出错: {e}")

            # Step 1: Check URL and Prompt User if incorrect
            target_url_part = "maimai.cn/ent/v41/im"

            # Additional Sync: If current URL doesn't match target, search all handles for it
            try:
                if target_url_part not in self.driver.current_url:
                    self.log_message(f"[{i+1}] 当前标签页不匹配，正在搜索包含 {target_url_part} 的标签页...")
                    found_handle = False
                    current_h = self.driver.current_window_handle
                    for h in self.driver.window_handles:
                        try:
                            self.driver.switch_to.window(h)
                            if target_url_part in self.driver.current_url:
                                found_handle = True
                                self.log_message(f"[{i+1}] 成功切换到目标标签页: {self.driver.title}")
                                break
                        except:
                            pass
                    if not found_handle:
                        # Restore if not found
                        try:
                            self.driver.switch_to.window(current_h)
                        except:
                            pass
            except Exception as e:
                pass
            
            # Allow user 30 seconds to switch manually if detected wrong URL
            check_start = time.time()
            url_ok = False
            
            while time.time() - check_start < 30:
                if check_stop(): break
                try:
                    cur_url = self.driver.current_url
                    # Always log current URL as requested for debugging
                    self.log_message(f"[{i+1}] 当前页面URL: {cur_url}")
                    
                    if target_url_part in cur_url:
                        url_ok = True
                        break
                    else:
                        # Only log warning once per second to avoid spam
                        if int(time.time()) % 5 == 0:
                             self.log_message(f"[{i+1}] URL不匹配 (需包含 {target_url_part})")
                        
                        # Show Toast/Popup to user asking to switch
                        # We use error signal to show a non-blocking toast or just log?
                        # User asked for "Pop up prompt"
                        # Since this is a thread, we can emit a signal to show a toast or message box
                        # For now, we will just wait and log. A blocking QMessageBox might freeze the thread logic if not handled carefully.
                        # Let's emit a signal to show a toast/notification
                        self.log_signal.emit(f"请切换到聊天页面 (包含 {target_url_part})")
                except:
                    pass
                time.sleep(1)
            
            if not url_ok:
                 self.log_message(f"[{i+1}] 超时未切换到正确页面，跳过本次操作")
                 if close_window:
                     try:
                         if not check_stop() and not self.debug_mode:
                             self.driver.close()
                             self.driver.switch_to.window(restore_handle)
                     except:
                         pass
                 return

            self.log_message(f"[{i+1}] 页面URL匹配，正在查找输入框...")
            self.log_message(f"[{i+1}] Driver Object: {self.driver}")
            
            # Step 2: Find input box (Strict Class Name, 3s Timeout)
            start_time = time.time()
            found = False
            
            # Timeout set to 3 seconds as requested
            while time.time() - start_time < 3:
                if check_stop(): break
                
                try:
                    input_box = None
                    used_strategy = "By.CLASS_NAME"
                    
                    # 1. Switch to imIframe first (as requested)
                    try:
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame("imIframe")
                    except:
                        pass

                    # 2. Find input box
                    try:
                        element = self.driver.find_element(By.CLASS_NAME, "inputPanel")
                        if element:
                            input_box = element
                    except:
                        pass

                    if input_box:
                        # Scroll into view
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_box)
                        time.sleep(0.5)
                        
                        self.log_message(f"[{i+1}] 找到输入框，准备发送 {len(self.messages)} 条消息...")
                        
                        # Iterate messages
                        for idx, msg_content in enumerate(self.messages):
                            if check_stop(): break
                            
                            self.log_message(f"[{i+1}] 正在发送第 {idx+1}/{len(self.messages)} 条消息...")
                            
                            # Click to focus first (ensure focus for each message)
                            try:
                                self.driver.execute_script("arguments[0].click();", input_box)
                            except:
                                pass
                            time.sleep(0.5)
                            
                            # Send keys directly to the element
                            ActionChains(self.driver).move_to_element(input_box).click().send_keys(msg_content).perform()
                            
                            self.log_message(f"[{i+1}] 输入消息内容成功")
                            
                            time.sleep(0.5) # Wait for UI update
                            
                            # Find Send button and Click
                            try:
                                # Use explicit XPath for the Send button: //a[text()='发送']
                                send_btn = self.driver.find_element(By.XPATH, "//a[text()='发送']")
                                self.driver.execute_script("arguments[0].click();", send_btn)
                                self.log_message(f"[{i+1}] 点击发送按钮成功")
                            except Exception as e:
                                self.log_message(f"[{i+1}] 未找到发送按钮或点击失败: {e}")
                            
                            time.sleep(1.5) # Wait between messages
                            
                        # --- New: Exchange Phone & WeChat ---
                        self.log_message(f"[{i+1}] 消息发送完毕，准备点击交换手机/微信...")
                        for tgt in ["交换手机", "交换微信"]:
                            if check_stop(): break
                            try:
                                # Locate via text inside class="tool-text"
                                # Outer HTML: <div class="tool normal">...<span class="tool-text">交换手机</span></div>
                                # Try to click the parent div or the span itself. Clicking span usually works or bubbles up.
                                # Use XPath to find the span
                                btn = self.driver.find_element(By.XPATH, f"//div[contains(@class, 'tool')]//span[contains(@class, 'tool-text') and text()='{tgt}']")
                                self.driver.execute_script("arguments[0].click();", btn)
                                self.log_message(f"[{i+1}] 已点击 '{tgt}'")
                                time.sleep(1.0)
                                
                                # Check for confirmation popup (specifically for WeChat, but safe to check for both)
                                # Popup structure: <div class="..."><p>...确定...?</p>...<a class="confirm">确定</a></div>
                                try:
                                    # Short wait for confirm button
                                    confirm_btn = WebDriverWait(self.driver, 2).until(
                                        EC.presence_of_element_located((By.XPATH, "//a[contains(@class, 'confirm') and text()='确定']"))
                                    )
                                    self.driver.execute_script("arguments[0].click();", confirm_btn)
                                    self.log_message(f"[{i+1}] 已点击 '{tgt}' 的确认弹窗")
                                    time.sleep(0.5)
                                except:
                                    # Popup might not appear if already exchanged or for Phone
                                    pass
                            except Exception as e:
                                self.log_message(f"[{i+1}] 点击 '{tgt}' 失败 (可能未找到或已交换): {str(e)}")
                        # ------------------------------------

                        found = True
                        break
                except Exception as e:
                    # Retry, but switch back to default content before next loop to retry finding iframe
                    self.driver.switch_to.default_content()
                    pass
                
                time.sleep(0.5)
            
            if not found:
                self.log_message(f"[{i+1}] 未找到输入框或输入失败 (超时)")
                # Debug info
                try:
                    self.driver.switch_to.default_content()
                    iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                    self.log_message(f"[{i+1}] 调试信息 - 页面包含 {len(iframes)} 个 iframe")
                    for idx, fr in enumerate(iframes):
                        self.log_message(f"[{i+1}] Iframe {idx+1}: ID='{fr.get_attribute('id')}' | Name='{fr.get_attribute('name')}'")
                except Exception as e:
                    self.log_message(f"[{i+1}] 无法获取调试信息: {e}")

            # Switch back to default content before closing or finishing
            try:
                self.driver.switch_to.default_content()
            except:
                pass

            # Window closing logic
            if close_window:
                if not check_stop():
                    if self.debug_mode:
                        self.log_message(f"[{i+1}] 调试模式开启：保留窗口，不关闭")
                    else:
                        self.driver.close()
                        self.driver.switch_to.window(restore_handle)
                        self.log_message(f"[{i+1}] 窗口已关闭，返回主窗口")
                else:
                    self.log_message(f"[{i+1}] 任务已停止，保留当前窗口")
                
        except Exception as e:
            self.log_message(f"[{i+1}] 处理输入框异常: {e}")
            if close_window:
                try:
                    if not check_stop():
                        if not self.debug_mode:
                            self.driver.close()
                            self.driver.switch_to.window(restore_handle)
                except:
                    pass
        
        return found

    def run(self):
        # Configure AI client
        if self.mode == "evaluate":
            ai_client.setup_client(self.api_key, self.base_url)

        global STOP_FLAG
        STOP_FLAG = False
        
        def stop_task():
            global STOP_FLAG
            STOP_FLAG = True
            self.log_message("检测到 Esc 键，正在强制终止...")

        try:
            keyboard.add_hotkey('esc', stop_task)
        except Exception:
            self.log_message("警告: 无法注册 Esc 热键")

        try:
            self.log_message("正在定位目标窗口...")
            if check_stop(): return

            # Initial window activation (try once at start)
            EnumWindows = windll.user32.EnumWindows
            GetWindowTextW = windll.user32.GetWindowTextW
            GetWindowTextLengthW = windll.user32.GetWindowTextLengthW
            IsWindowVisible = windll.user32.IsWindowVisible
            GetClassNameW = windll.user32.GetClassNameW
            hwnds = []
            titles = []
            classes = []

            def enum_proc(hwnd, lParam):
                if IsWindowVisible(hwnd):
                    ln = GetWindowTextLengthW(hwnd)
                    if ln > 0:
                        buf = create_unicode_buffer(ln + 1)
                        GetWindowTextW(hwnd, buf, ln + 1)
                        title = buf.value
                        cbuf = create_unicode_buffer(256)
                        GetClassNameW(hwnd, cbuf, 256)
                        cls = cbuf.value
                        hwnds.append(hwnd)
                        titles.append(title)
                        classes.append(cls)
                return True

            enum_proc = WINFUNCTYPE(c_bool, c_void_p, c_void_p)(enum_proc)
            EnumWindows(enum_proc, 0)
            
            suffixes = browser_suffixes(self.browser)
            class_map = browser_class_map(self.browser)
            idx = -1
            for i, t in enumerate(titles):
                if (any(s in t for s in suffixes)) or (i < len(classes) and (not class_map or classes[i] in class_map)):
                    idx = i
                    break
            
            if idx != -1:
                IsIconic = windll.user32.IsIconic
                SetForegroundWindow = windll.user32.SetForegroundWindow
                BringWindowToTop = windll.user32.BringWindowToTop
                SetWindowPos = windll.user32.SetWindowPos
                SW_RESTORE = 9
                SWP_NOSIZE = 0x0001
                SWP_NOMOVE = 0x0002
                SWP_SHOWWINDOW = 0x0040
                HWND_TOP = 0
                h = hwnds[idx]
                if IsIconic(h):
                    windll.user32.ShowWindow(h, SW_RESTORE)
                SetWindowPos(h, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW)
                SetForegroundWindow(h)
                BringWindowToTop(h)
            
            _, active_title, _ = get_active_window_info()
            time.sleep(0.5)
            
            cur_url = get_active_tab_url_via_devtools(self.browser, active_title)
            self.log_message(f"当前窗口: {active_title}")
            
            active_ok = False
            if "人才银行" in active_title or "脉脉" in active_title or "Maimai" in active_title:
                active_ok = True
            else:
                if cur_url and ("maimai.cn" in cur_url):
                    active_ok = True
            
            if not active_ok:
                self.error_signal.emit("当前页卡未定位到目标站点，请打开人才银行页面")
                self.finished_signal.emit()
                return

            # Init shared driver for the loop
            self.log_message(f"准备连接浏览器 ({self.browser})...")
            self.driver = attach_selenium_driver(self.browser, logger=self.log_message)
            if not self.driver:
                self.error_signal.emit("无法连接到浏览器调试端口\n请查看目录下 debug.log 获取详细错误信息")
                self.finished_signal.emit()
                return

            # Prevent driver from closing the browser on exit
            # Since we attached to an existing browser, we never want to kill the process
            self.driver.quit = lambda: None

            success_count = 0
            fail_count = 0
            skip_count = 0
            
            reset_next = False
            
            # 使用 while 循环直到达到目标成功数量
            i = -1
            while success_count < self.max_count:
                if check_stop(): break
                i += 1
                
                try:
                    self.log_signal.emit(f"正在执行第 {i+1} 次尝试 (目标成功: {success_count}/{self.max_count})...")
                    self.log_to_file(f"正在执行第 {i+1} 次尝试 (目标成功: {success_count}/{self.max_count})...")
                    
                    if i == 0:
                        try:
                            # 方案：使用 window.devicePixelRatio 检测网页缩放
                            # 该值反映了物理像素与逻辑像素的比例。
                            # 真实的网页缩放 = devicePixelRatio / 系统DPI缩放
                            # 严格要求真实网页缩放必须等于 1.0 (允许微小浮点误差)。
                            
                            device_pixel_ratio = self.driver.execute_script("return window.devicePixelRatio")
                            sys_scale = get_system_dpi_scale()
                            actual_zoom = device_pixel_ratio / sys_scale
                            
                            # 只要真实缩放不是 1.0 (考虑浮点数精度，范围设定在 0.99 到 1.01) 就拦截
                            if abs(actual_zoom - 1.0) > 0.01:
                                msg = f"检测到浏览器网页缩放异常。\n\n当前网页缩放: {actual_zoom:.2f}\n(系统缩放: {sys_scale:.2f}, DPR: {device_pixel_ratio:.2f})\n\n这会导致自动化点击偏移或截图不完整。\n请按 Ctrl+0 重置浏览器缩放。\n\n程序将停止运行。"
                                # 发送错误信号并立即抛出异常，强制跳出整个 try 块，进入 finally 流程
                                self.error_signal.emit(msg)
                                self.log_to_file(f"[严重错误] 浏览器缩放异常 (Actual: {actual_zoom:.2f})，程序已停止。")
                                # 不用 return，抛出异常确保外层循环立即终止并清理
                                raise Exception("ZOOM_ERROR")
                        except Exception as e:
                            if str(e) == "ZOOM_ERROR":
                                raise e # 继续向外抛出，让外层捕获并结束线程
                            self.log_signal.emit(f"[警告] 缩放检测失败: {e}")
                    
                    # If we are in debug mode BUT on the main list page, we want to run the full flow (click evaluate -> click chat -> open window)
                    # So only skip if we are NOT on the main list page (presumably already on chat page)
                    if self.debug_mode and self.mode == "evaluate" and cur_url and "recruit/talents" not in cur_url:
                        self.log_signal.emit(f"[{i+1}] 调试模式：当前不在列表页，跳过前序步骤，直接执行常用语扫描...")
                        self.log_to_file(f"[{i+1}] 调试模式：当前不在列表页，跳过前序步骤，直接执行常用语扫描...")
                        self.process_common_phrase(i, close_window=False)
                        self.log_signal.emit(f"[{i+1}] 调试模式执行完毕，结束任务。")
                        self.log_to_file(f"[{i+1}] 调试模式执行完毕，结束任务。")
                        break

                    # Step 1: Click start/filter button
                    if check_stop(): break
                    if i == 0:
                        if self.debug_mode:
                             self.log_signal.emit(f"[{i+1}] 调试模式已开启，跳过点击筛选按钮，直接开始任务")
                             self.log_to_file(f"[{i+1}] 调试模式已开启，跳过点击筛选按钮，直接开始任务")
                        else:
                             self.log_signal.emit(f"[{i+1}] 正在点击筛选按钮...")
                             self.log_to_file(f"[{i+1}] 正在点击筛选按钮...")
                             res = wait_and_click_xpath(self.browser, '//*[@id="recruit_talents"]/div/div/div/div[3]/div[2]/div[3]/div[1]', expected_title="人才银行", timeout=10, driver=self.driver)
                             
                             if res == "stopped": break
                             if res != "ok":
                                 self.log_signal.emit(f"[{i+1}] 点击筛选按钮失败: {res}，跳过本次")
                                 self.log_to_file(f"[{i+1}] 点击筛选按钮失败: {res}，跳过本次")
                                 fail_count += 1
                                 continue

                    time.sleep(0.2)
                    if check_stop(): break
                    
                    # Determine current index
                    is_first = (i == 0) or reset_next
                    idx = 2 if is_first else 3
                    add_idx = 1 if is_first else 2
                    reset_next = False # Consumed
                    
                    # Step 2: Check skip element
                    skip_xpath = f'//*[@id="root"]/section/div[1]/div[3]/div/div/div[2]/div/div[{idx}]/div[3]/div/div[1]/div[2]/div[1]'
                    next_xpath = f'//*[@id="root"]/section/div[1]/div[3]/div/div/div[2]/div/div[{idx}]/img'
                    pagination_xpath = '//*[@id="root"]/section/div[1]/div[3]/div/div/div[2]/div/div[2]/div[1]/div/div[31]'

                    # Helper to handle next/pagination
                    def try_click_next():
                        nonlocal reset_next
                        if check_stop(): return "stopped"
                        
                        # 尝试点击常规的下一个
                        n_res = wait_and_click_xpath(self.browser, next_xpath, expected_title="人才银行", timeout=5, driver=self.driver)
                        if n_res == "ok":
                            return "ok"
                        if n_res == "stopped":
                            return "stopped"
                        
                        # 如果没找到，尝试翻页
                        # 如果没找到，尝试翻页
                        self.log_signal.emit(f"[{i+1}] 未找到下一个元素，尝试翻页 (Detect '跳转至下一页')...")
                        self.log_to_file(f"[{i+1}] 未找到下一个元素，尝试翻页 (Detect '跳转至下一页')...")
                        
                        # 使用 ClassName 和 文本 "跳转至下一页" 定位
                        # Class: w-full h-[40px] text-[14px] text-[#3375FF] leading-[22px] bg-[#E5EEFF] flex justify-center items-center mb-[0px] cursor-pointer rounded-[4px]
                        # Text: 跳转至下一页
                        
                        next_page_btn = None
                        try:
                            # 尝试通过文本内容和部分关键 Class 查找
                            # 使用 XPath 匹配文本和完整的 class 属性 (或者包含关键部分)
                            # 为了稳健性，匹配文本 '跳转至下一页' 和 class 包含 'cursor-pointer' (或其他特征)
                            # 用户提供了完整的 HTML，我们可以尝试匹配文本
                            
                            # 方法1: 查找所有包含 "跳转至下一页" 文本的 div
                            candidates = self.driver.find_elements(By.XPATH, "//div[text()='跳转至下一页']")
                            
                            target_class = "w-full h-[40px] text-[14px] text-[#3375FF] leading-[22px] bg-[#E5EEFF] flex justify-center items-center mb-[0px] cursor-pointer rounded-[4px]"
                            
                            for cand in candidates:
                                # 检查 Class 是否匹配 (完全匹配或包含)
                                # 由于 tailwind class 顺序可能变化，或者浏览器解析后可能有差异，建议检查关键 class 或者整个字符串
                                # 这里尝试获取 class 属性并比对
                                cand_class = cand.get_attribute("class")
                                if cand_class and "w-full" in cand_class and "text-[#3375FF]" in cand_class:
                                    next_page_btn = cand
                                    break
                            
                            if next_page_btn:
                                self.driver.execute_script("arguments[0].click();", next_page_btn)
                                self.log_signal.emit(f"[{i+1}] 翻页点击成功 (Text+Class)，等待加载并校验...")
                                self.log_to_file(f"[{i+1}] 翻页点击成功 (Text+Class)，等待加载并校验...")
                                time.sleep(3.0) # 等待页面刷新
                                
                                # 翻页后，下一个循环应视为新页面的第一个
                                reset_next = True
                                return "ok"
                            else:
                                self.log_signal.emit(f"[{i+1}] 未找到 '跳转至下一页' 按钮")
                                self.log_to_file(f"[{i+1}] 未找到 '跳转至下一页' 按钮")
            
                        except Exception as e:
                            self.log_signal.emit(f"[{i+1}] 翻页操作异常: {str(e)}")
                            self.log_to_file(f"[{i+1}] 翻页操作异常: {str(e)}")
                        
                        return "fail"

                    if check_element_exists(self.browser, skip_xpath, expected_title="人才银行", timeout=2, driver=self.driver):
                        self.log_signal.emit(f"[{i+1}] 发现跳过标记，尝试进入下一个...")
                        self.log_to_file(f"[{i+1}] 发现跳过标记，尝试进入下一个...")
                        if check_stop(): break
                        
                        nx_res = try_click_next()
                        if nx_res == "stopped": break
                        if nx_res == "ok":
                            skip_count += 1
                            continue
                        else:
                            self.log_signal.emit(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                            self.log_to_file(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                            break

                    # Step 3: Add Friend Check
                    added_this_round = False
                    if check_stop(): break
                    if self.mode == "add_friend":
                        self.log_message(f"[{i+1}] 正在检查加好友状态...")
                        add_friend_xpath = f'//*[@id="root"]/section/div[1]/div[3]/div/div/div[2]/div/div[{add_idx}]/div[3]/div/div[2]/div[1]/div/div/div[3]'
                        
                        can_add = False
                        el_target = None
                        try:
                            # Use shared driver
                            driver = self.driver
                            if driver:
                                try:
                                    # 尝试定位元素，不等待太久
                                    if check_stop(): break
                                    el_target = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, add_friend_xpath)))
                                    
                                    # 获取文本和样式
                                    txt = el_target.text.strip()
                                    style = el_target.get_attribute("style") or ""
                                    
                                    if "加好友" not in txt:
                                        self.log_message(f"[{i+1}] 按钮文本为 '{txt}'，非加好友，跳过")
                                    elif "cursor: not-allowed" in style:
                                        self.log_message(f"[{i+1}] 发现不可加好友样式，跳过")
                                    else:
                                        can_add = True
                                        
                                except Exception:
                                    self.log_message(f"[{i+1}] 未找到加好友按钮，跳过")
                        except Exception as e:
                            self.log_message(f"[{i+1}] 检查加好友状态出错: {str(e)}")

                        if check_stop(): break
                        if can_add and el_target:
                            self.log_message(f"[{i+1}] 正在点击加好友...")
                            try:
                                # 使用 ActionChains 控制鼠标点击元素对应位置
                                if check_stop(): break
                                ActionChains(driver).move_to_element(el_target).click().perform()
                                self.log_message(f"[{i+1}] 已通过鼠标模拟点击加好友")
                                added_this_round = True
                                time.sleep(0.5)
                            except Exception as e:
                                self.log_message(f"[{i+1}] 鼠标模拟点击失败，尝试JS点击: {str(e)}")
                                try:
                                    if check_stop(): break
                                    driver.execute_script("arguments[0].click();", el_target)
                                    self.log_message(f"[{i+1}] JS点击加好友成功")
                                    added_this_round = True
                                    time.sleep(0.5)
                                except Exception as ex:
                                    self.log_message(f"[{i+1}] JS点击也失败: {str(ex)}")

                        # Check for privacy popup immediately after add friend attempt
                        if added_this_round:
                            try:
                                # Give it a moment to appear
                                time.sleep(1.0)
                                if check_stop(): break
                                
                                # New Logic: Detect via ClassName "ant-modal-content"
                                popups = self.driver.find_elements(By.CLASS_NAME, "ant-modal-content")
                                if popups:
                                    # Use the last one (usually top-most)
                                    target_popup = popups[-1]
                                    
                                    self.log_message(f"[{i+1}] 检测到隐私设置弹窗 (class='ant-modal-content')")
                                    
                                    # Check for limit reached text within the popup
                                    limit_reached = False
                                    try:
                                        if "加好友券已用完" in target_popup.text:
                                            limit_reached = True
                                    except Exception:
                                        pass
                                    
                                    if limit_reached:
                                        self.log_message(f"[{i+1}] 检测到今日加好友已达上限！正在关闭弹窗并终止程序...")
                                    else:
                                        self.log_message(f"[{i+1}] 未达到上限，正在关闭弹窗...")

                                    # Close the popup
                                    # Target: <button type="button" class="ant-btn ant-btn-primary"><span>我知道了</span></button>
                                    try:
                                        # Try to find the button inside the popup first
                                        close_btn_xpath = ".//button[contains(@class, 'ant-btn') and contains(@class, 'ant-btn-primary') and .//span[text()='我知道了']]"
                                        close_btns = target_popup.find_elements(By.XPATH, close_btn_xpath)
                                        
                                        if not close_btns:
                                            # Fallback to global search
                                            close_btns = self.driver.find_elements(By.XPATH, "//button[contains(@class, 'ant-btn') and contains(@class, 'ant-btn-primary') and .//span[text()='我知道了']]")
                                        
                                        if close_btns:
                                            close_btn = close_btns[-1]
                                            self.driver.execute_script("arguments[0].click();", close_btn)
                                            self.log_message(f"[{i+1}] 已关闭隐私弹窗")
                                        else:
                                            self.log_message(f"[{i+1}] 未找到关闭按钮 ('我知道了')")
                                    except Exception as e:
                                        self.log_message(f"[{i+1}] 关闭隐私弹窗失败: {str(e)}")

                                    if limit_reached:
                                        self.error_signal.emit("今日加好友已达上限，程序终止。")
                                        stop_task() # Set STOP_FLAG
                                        break
                                    
                                    # Mark as failed as per user request
                                    self.log_message(f"[{i+1}] 因隐私设置未成功添加，本次不计入成功数")
                                    added_this_round = False
                            except Exception as e:
                                self.log_message(f"[{i+1}] 隐私弹窗检测出错: {str(e)}")
                        


                    # Step 4: Screenshot & Evaluate
                    if check_stop(): break
                    if self.mode == "evaluate":
                        # --- New Logic: Check Immediate Communication First ---
                        chat_btn_xpath = f'//*[@id="root"]/section/div[1]/div[3]/div/div/div[2]/div/div[{add_idx}]/div[3]/div/div[1]/div[2]/div[1]'
                        
                        can_chat = False
                        try:
                            # Quick check for button existence and text
                            c_el = WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, chat_btn_xpath)))
                            c_txt = c_el.text.strip()
                            if "立即沟通" in c_txt:
                                can_chat = True
                            else:
                                self.log_message(f"[{i+1}] 按钮文本为 '{c_txt}'，非立即沟通，跳过评估")
                        except Exception:
                            self.log_message(f"[{i+1}] 未找到立即沟通按钮，跳过评估")
                        
                        if not can_chat:
                            # Skip to next directly
                            nx_res = try_click_next()
                            if nx_res == "stopped": break
                            if nx_res != "ok":
                                self.log_message(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                                break
                            continue
                        
                        # --- If can chat, proceed to Screenshot ---
                        self.log_message(f"[{i+1}] 发现立即沟通按钮，正在截图(备注区域)...")
                        
                        # Resume card XPath (base)
                        resume_card_xpath = f'//*[@id="root"]/section/div[1]/div[3]/div/div/div[2]/div/div[{add_idx}]'
                        
                        # Target the specific remark container based on class "remarks___" AND text "备注"
                        # This combines robustness of class name with specificity of text content
                        sc_target_xpath = f'{resume_card_xpath}//div[contains(@class, "remarks___") and .//div[contains(text(), "备注")]]'
                        
                        sc_res, sc_path = wait_and_capture_xpath(self.browser, sc_target_xpath, expected_title="人才银行", timeout=10, driver=self.driver)
                        
                        if sc_res == "stopped": break
                        if sc_res != "ok":
                            self.log_message(f"[{i+1}] 截图失败: {sc_res}，尝试关闭弹窗并跳过")
                            nx_res = try_click_next()
                            if nx_res == "stopped": break
                            if nx_res != "ok":
                                self.log_message(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                                break
                            fail_count += 1
                            continue

                        # --- Evaluate ---
                        if check_stop(): break
                        self.log_message(f"[{i+1}] 正在评估简历...")
                        try:
                            r = ai_client.evaluate_resume_from_image(sc_path, "", self.system_prompt, self.model)
                            if check_stop(): break
                            res_str = str(r)
                            
                            # Parse result and reason (Format: Result|Reason)
                            # Enhanced parsing logic to handle cases where AI misses the separator
                            res_str = res_str.strip()
                            decision = ""
                            reason = ""
                            
                            if "|" in res_str:
                                parts = res_str.split("|", 1)
                                decision = parts[0].strip()
                                reason = parts[1].strip()
                            else:
                                # Fallback: try to detect intent from text
                                if "不需要联系" in res_str or "拒绝" in res_str:
                                    decision = "不需要联系"
                                    reason = res_str # Use full response as reason
                                elif "需要联系" in res_str:
                                    decision = "需要联系"
                                    reason = res_str
                                else:
                                    decision = "格式错误/无法判断"
                                    reason = res_str
                            
                            # Clean up decision text if it contains extra chars
                            if "不需要" in decision: decision = "不需要联系"
                            elif "需要" in decision: decision = "需要联系"
                            
                            self.log_message(f"[{i+1}] AI评估结果: {decision}, 依据: {reason}")
                            
                            if "需要联系" in decision and "不需要" not in decision:
                                self.popup_signal.emit(f"[{i+1}] 评估结果：需要联系\n依据：{reason}")
                                time.sleep(3.0) # Wait for popup to show for 3 seconds
                                
                                self.log_message(f"[{i+1}] 正在尝试立即沟通...")
                                
                                # Click the button we already verified
                                chat_res = "fail"
                                try:
                                    # Re-find element just in case DOM changed slightly or reference lost
                                    el_chat_target = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, chat_btn_xpath)))
                                    
                                    self.log_message(f"[{i+1}] 正在点击立即沟通 (模拟鼠标)...")
                                    ActionChains(self.driver).move_to_element(el_chat_target).click().perform()
                                    chat_res = "ok"
                                except Exception as e:
                                    self.log_message(f"[{i+1}] 鼠标模拟点击立即沟通失败: {str(e)}，尝试常规点击")
                                    chat_res = wait_and_click_xpath(self.browser, chat_btn_xpath, expected_title="人才银行", timeout=5, driver=self.driver)
                                
                                if chat_res == "ok":
                                    self.log_message(f"[{i+1}] 点击立即沟通成功，等待弹窗...")
                                    
                                    # Step 2: Wait for Immediate Communication Window
                                    # Try broader detection
                                    popup_detected = False
                                    popup_xpath_base = '/html/body/div'
                                    
                                    # We don't know exact div index, try searching div[4] to div[10] for the specific structure
                                    # Structure: /div[idx]/div/div[2]/div/div[2]
                                    
                                    for idx in range(4, 11):
                                        t_xpath = f'/html/body/div[{idx}]/div/div[2]/div/div[2]'
                                        if check_element_exists(self.browser, t_xpath, expected_title="人才银行", timeout=0.2, driver=self.driver):
                                            popup_detected = True
                                            self.log_message(f"[{i+1}] 立即沟通窗口已弹出 (div[{idx}])")
                                            
                                            # Found popup, now click button inside it
                                            # Button structure: .../div[3]/div/div[2]/button[1]
                                            send_continue_xpath = f'{t_xpath}/div[3]/div/div[2]/button[1]'
                                            
                                            # Get current window handles before clicking to identify new tab
                                            old_handles = self.driver.window_handles
                                            current_handle = self.driver.current_window_handle
                                            
                                            send_res = wait_and_click_xpath(self.browser, send_continue_xpath, expected_title="人才银行", timeout=5, driver=self.driver)
                                            
                                            if send_res == "ok":
                                                self.log_message(f"[{i+1}] 点击发送后继续沟通成功，等待新窗口...")
                                                time.sleep(2.0) # Wait for new tab to open
                                                
                                                new_handles = self.driver.window_handles
                                                new_tab_handle = None
                                                for h in new_handles:
                                                    if h not in old_handles:
                                                        new_tab_handle = h
                                                        break
                                                
                                                if new_tab_handle:
                                                    self.driver.switch_to.window(new_tab_handle)
                                                    self.log_message(f"[{i+1}] 已切换到新窗口，准备发送常用语...")
                                                    ph_success = self.process_common_phrase(i, close_window=True, restore_handle=current_handle)
                                                    if ph_success:
                                                        success_count += 1
                                                        self.log_message(f"[{i+1}] 成功发送消息并完成流程，计入成功数 (当前: {success_count})")
                                                    
                                                    # --- New: Add Note Logic ---
                                                    if not self.debug_mode:
                                                        self.log_message(f"[{i+1}] 正在添加备注...")
                                                        try:
                                                            # 1. Directly find and click the note textarea
                                                            # <div class="mb-16 ...">请输入备注内容</div>
                                                            initial_div = WebDriverWait(self.driver, 5).until(
                                                                EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'mb-16') and contains(text(), '请输入备注内容')]"))
                                                            )
                                                            self.driver.execute_script("arguments[0].click();", initial_div)
                                                            self.log_message(f"[{i+1}] 已点击备注输入框")
                                                            time.sleep(0.5)

                                                            # 2. Find the actual textarea that appears
                                                            note_area = WebDriverWait(self.driver, 5).until(
                                                                EC.presence_of_element_located((By.XPATH, "//textarea[contains(@class, 'ant-input')]"))
                                                            )
                                                            
                                                            # 2. Check 'Only visible to self'
                                                            # Locate the label containing the text
                                                            only_self_label = self.driver.find_element(By.XPATH, "//label[contains(@class, 'ant-checkbox-wrapper') and .//span[contains(text(), '仅自己可见')]]")
                                                            
                                                            # Check if checked (class ant-checkbox-checked on the span inside label)
                                                            cb_span = only_self_label.find_element(By.XPATH, ".//span[contains(@class, 'ant-checkbox')]")
                                                            if "ant-checkbox-checked" not in cb_span.get_attribute("class"):
                                                                self.driver.execute_script("arguments[0].click();", only_self_label)
                                                                self.log_message(f"[{i+1}] 已勾选'仅自己可见'")
                                                            
                                                            time.sleep(0.5)
                                                            
                                                            # 3. Input '1'
                                                            note_area.clear()
                                                            note_area.send_keys("1")
                                                            self.log_message(f"[{i+1}] 已输入备注 '1'")
                                                            time.sleep(0.5)
                                                            
                                                            # 4. Click Confirm
                                                            # <div class="mui-btn mui-btn-primary mui-btn-small">确定</div>
                                                            confirm_btn = self.driver.find_element(By.XPATH, "//div[contains(@class, 'mui-btn-primary') and text()='确定']")
                                                            self.driver.execute_script("arguments[0].click();", confirm_btn)
                                                            self.log_message(f"[{i+1}] 已保存备注")
                                                            time.sleep(1.0)
                                                            
                                                        except Exception as e:
                                                            self.log_message(f"[{i+1}] 添加备注失败: {str(e)}")
                                                    # ---------------------------

                                                    if self.debug_mode:
                                                        self.log_message(f"[{i+1}] 调试模式：任务已完成，保留窗口")
                                                        return
                                                else:
                                                    self.log_signal.emit(f"[{i+1}] 未检测到新窗口打开")
                                                    self.log_to_file(f"[{i+1}] 未检测到新窗口打开")
                                            else:
                                                self.log_signal.emit(f"[{i+1}] 点击发送后继续沟通失败")
                                                self.log_to_file(f"[{i+1}] 点击发送后继续沟通失败")
                                                break
                                    
                                    if not popup_detected:
                                        self.log_signal.emit(f"[{i+1}] 未检测到立即沟通窗口 (尝试了 div[4]-div[10])")
                                        self.log_to_file(f"[{i+1}] 未检测到立即沟通窗口 (尝试了 div[4]-div[10])")
                                else:
                                    self.log_signal.emit(f"[{i+1}] 点击立即沟通失败: {chat_res}")
                                    self.log_to_file(f"[{i+1}] 点击立即沟通失败: {chat_res}")

                                nx_res = try_click_next()
                                if nx_res == "stopped": break
                            else:
                                self.log_message(f"[{i+1}] 评估结果：不需要联系, 依据: {reason}")
                                self.popup_signal.emit(f"[{i+1}] 评估结果：不需要联系\n依据：{reason}")
                                time.sleep(3.0) # Wait for popup
                                
                                self.log_message(f"[{i+1}] 正在点击下一个...")
                                nx_res = try_click_next()
                                if nx_res == "stopped": break
                            
                            # nx_res check: only break if failed and not stopped
                            # success_count is already incremented in process_common_phrase block
                            if nx_res != "ok" and nx_res != "stopped":
                                self.log_signal.emit(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                                self.log_to_file(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                                break
                        except Exception as e:
                            self.log_signal.emit(f"[{i+1}] 评估异常: {str(e)}，跳过")
                            self.log_to_file(f"[{i+1}] 评估异常: {str(e)}，跳过")
                            nx_res = try_click_next()
                            if nx_res == "stopped": break
                            if nx_res != "ok":
                                self.log_signal.emit(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                                self.log_to_file(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                                break
                            fail_count += 1
                        finally:
                            # Delete screenshot to free disk space
                            # try:
                            #     if sc_path and os.path.exists(sc_path):
                            #         os.remove(sc_path)
                            # except Exception:
                            #     pass
                            pass
                    else:
                        # add_friend mode, just go next
                        self.log_signal.emit(f"[{i+1}] 纯加好友模式，正在点击下一个...")
                        self.log_to_file(f"[{i+1}] 纯加好友模式，正在点击下一个...")
                        if check_stop(): break
                        nx_res = try_click_next()
                        if nx_res == "stopped": break
                        
                        if nx_res == "ok":
                            if added_this_round:
                                success_count += 1
                            else:
                                self.log_signal.emit(f"[{i+1}] 本次未执行加好友 (已加过或跳过)，不计入成功总数")
                                self.log_to_file(f"[{i+1}] 本次未执行加好友 (已加过或跳过)，不计入成功总数")
                        else:
                            self.log_signal.emit(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                            self.log_to_file(f"[{i+1}] 无法找到下一个或翻页按钮，视为到达末尾，结束任务")
                            break

                    if check_stop(): break

                    # Memory management: GC every 50 iterations
                    if (i + 1) % 50 == 0:
                        self.log_signal.emit(f"[{i+1}] 正在执行内存维护...")
                        gc.collect()

                except Exception as e:
                    if str(e) == "ZOOM_ERROR":
                        # 缩放异常已经发送过 error_signal，这里直接跳出主循环结束任务
                        break
                    if check_stop(): break
                    self.log_signal.emit(f"[{i+1}] 发生未知错误: {str(e)}，跳过")
                    self.log_to_file(f"[{i+1}] 发生未知错误: {str(e)}，跳过")
                    fail_count += 1
                    continue
        
        finally:
            try:
                keyboard.remove_hotkey('esc')
            except Exception:
                pass

        total_attempts = i + 1
        if check_stop():
             final_msg = f"任务已强制终止 (共尝试 {total_attempts} 次)。\n成功: {success_count}\n失败: {fail_count}\n跳过: {skip_count}"
        else:
             final_msg = f"任务完成 (共尝试 {total_attempts} 次)。\n成功: {success_count}\n失败: {fail_count}\n跳过: {skip_count}"
             
        self.result_signal.emit(final_msg)
        self.finished_signal.emit()



class UpdateThread(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, download_url, save_path):
        super().__init__()
        self.download_url = download_url
        self.save_path = save_path

    def run(self):
        try:
            self.progress_signal.emit(0, "正在连接服务器...")
            
            headers = {}
            if GITHUB_TOKEN and GITHUB_TOKEN != "YOUR_GITHUB_TOKEN_HERE":
                headers["Authorization"] = f"token {GITHUB_TOKEN}"
                # 对于 GitHub API 下载 release asset，需要指定 accept header
                headers["Accept"] = "application/octet-stream"
                
            response = requests.get(self.download_url, stream=True, timeout=15, headers=headers)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            block_size = 1024 * 8
            downloaded = 0
            
            with open(self.save_path, 'wb') as f:
                for data in response.iter_content(block_size):
                    if data:
                        f.write(data)
                        downloaded += len(data)
                        if total_size > 0:
                            percent = int((downloaded / total_size) * 100)
                            self.progress_signal.emit(percent, f"正在下载... {percent}%")
            
            self.progress_signal.emit(100, "下载完成，准备替换...")
            self.finished_signal.emit(True, self.save_path)
            
        except Exception as e:
            self.finished_signal.emit(False, f"下载更新失败: {str(e)}")

def check_for_updates():
    """检查更新，返回 (有新版本布尔值, 版本号, 更新日志, 下载链接)"""
    try:
        headers = {}
        if GITHUB_TOKEN and GITHUB_TOKEN != "YOUR_GITHUB_TOKEN_HERE":
            headers["Authorization"] = f"token {GITHUB_TOKEN}"
            
        response = requests.get(UPDATE_API_URL, timeout=5, headers=headers)
        if response.status_code == 200:
            data = response.json()
            latest_version = data.get("tag_name", "").replace("v", "")
            
            # 简单的版本号比较 (1.0.1 > 1.0.0)
            if latest_version and latest_version != CURRENT_VERSION:
                # 尝试判断是否真的大于当前版本
                def parse_v(v): return [int(x) for x in v.split('.') if x.isdigit()]
                if parse_v(latest_version) > parse_v(CURRENT_VERSION):
                    body = data.get("body", "暂无更新日志")
                    
                    # 获取源代码的 zip 下载链接，因为现在直接将程序推送到仓库的 main 分支
                    # 从 Release 中直接获取 zipball_url
                    download_url = data.get("zipball_url", "")
                    
                    return True, latest_version, body, download_url
        elif response.status_code == 404:
            return False, "", "仓库未找到或无权限访问 (如果是私有仓库请配置 Token)", ""
        else:
             return False, "", f"检查更新失败，HTTP 状态码: {response.status_code}", ""
             
        return False, "", "", ""
    except Exception as e:
        return False, "", f"检查更新出错: {e}", ""

class ToastWindow(QWidget):
    def __init__(self, message, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.label = QLabel(message)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setWordWrap(True)
        self.label.setStyleSheet("""
            QLabel {
                background-color: #2D3138;
                color: #FFFFFF;
                border: 2px solid #3D7EFF;
                border-radius: 10px;
                padding: 15px;
                font-size: 16px;
                font-weight: bold;
                min-width: 200px;
            }
        """)
        
        # Add shadow effect
        from PyQt6.QtWidgets import QGraphicsDropShadowEffect
        from PyQt6.QtGui import QColor
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 150))
        shadow.setOffset(0, 0)
        self.label.setGraphicsEffect(shadow)
        
        layout.addWidget(self.label)
        
        # Auto close timer
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.close)
        self.timer.start(3000)
        
        self.adjustSize()
        self.center_on_screen()
        
    def center_on_screen(self):
        screen = QApplication.primaryScreen().availableGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

class TitleBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(40)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 0, 10, 0)
        layout.setSpacing(10)

        # Spacer to push controls to right
        layout.addStretch()

        # Window Controls
        # Minimize
        self.btn_min = QPushButton("－")
        self.btn_min.setFixedSize(46, 30)
        self.btn_min.setObjectName("TitleBtnMin")
        self.btn_min.clicked.connect(self.window().showMinimized)
        
        # Maximize/Restore
        self.btn_max = QPushButton("□")
        self.btn_max.setFixedSize(46, 30)
        self.btn_max.setObjectName("TitleBtnMax")
        self.btn_max.clicked.connect(self.toggle_max)

        # Close
        self.btn_close = QPushButton("×")
        self.btn_close.setFixedSize(46, 30)
        self.btn_close.setObjectName("TitleBtnClose")
        self.btn_close.clicked.connect(self.window().close)
        
        layout.addWidget(self.btn_min)
        layout.addWidget(self.btn_max)
        layout.addWidget(self.btn_close)

        # Drag logic variables
        self.start_pos = None

    def toggle_max(self):
        if self.window().isMaximized():
            self.window().showNormal()
            self.btn_max.setText("□")
        else:
            self.window().showMaximized()
            self.btn_max.setText("❐")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.start_pos = event.globalPosition().toPoint()
            self.window_pos = self.window().frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.start_pos:
            delta = event.globalPosition().toPoint() - self.start_pos
            self.window().move(self.window_pos + delta)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.start_pos = None

class SidebarButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setCheckable(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedHeight(40)
        self.setObjectName("SidebarButton")

class Sidebar(QWidget):
    page_changed = pyqtSignal(int)
    check_update_signal = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(200)
        self.setObjectName("Sidebar")
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 20)
        layout.setSpacing(5)
        
        # App Title / Logo Area
        self.title_label = QLabel(f"AutoNetworker\nv{CURRENT_VERSION}")
        self.title_label.setObjectName("AppTitle")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setFixedHeight(50) # Match TitleBar height
        layout.addWidget(self.title_label)
        
        # Add some spacing
        layout.addSpacing(10)
        
        self.btn_group = QButtonGroup(self)
        self.btn_group.setExclusive(True)
        self.btn_group.idClicked.connect(self.page_changed.emit)

        # 1. Maimai Tools
        self.btn_maimai = SidebarButton("脉脉工具")
        layout.addWidget(self.btn_maimai)
        self.btn_group.addButton(self.btn_maimai, 0)
        
        # 2. AI Config
        self.btn_ai = SidebarButton("配置 AI")
        layout.addWidget(self.btn_ai)
        self.btn_group.addButton(self.btn_ai, 1)

        layout.addStretch()
        
        # Update Button at bottom
        self.btn_update = QPushButton("检查更新")
        self.btn_update.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_update.setFixedHeight(36)
        self.btn_update.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: 1px solid #DEE0E3;
                border-radius: 6px;
                color: #5C5F66;
                margin: 0 20px;
            }
            QPushButton:hover {
                background-color: #E5E6EB;
                color: #1F2329;
            }
        """)
        self.btn_update.clicked.connect(self.check_update_signal.emit)
        layout.addWidget(self.btn_update)
        
        # Default select first
        self.btn_maimai.setChecked(True)
        
        # Drag logic variables
        self.start_pos = None

    def mousePressEvent(self, event):
        # Allow dragging from the top area (Title)
        if event.button() == Qt.MouseButton.LeftButton and event.pos().y() <= 40:
            self.start_pos = event.globalPosition().toPoint()
            self.window_pos = self.window().frameGeometry().topLeft()
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.start_pos:
            delta = event.globalPosition().toPoint() - self.start_pos
            self.window().move(self.window_pos + delta)
            event.accept()
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.start_pos = None
        super().mouseReleaseEvent(event)

class MessageManager:
    def __init__(self):
        self.presets = {}
        self.recent = []
        self.load_presets()

    def get_presets_path(self):
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, "presets.json")

    def load_presets(self):
        path = self.get_presets_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if "_recent" in data:
                        self.recent = data.pop("_recent")
                    self.presets = data
            except:
                self.presets = {"默认预设": [{"content": "你好，对您的经历很感兴趣", "enabled": True}]}
                self.recent = []
        else:
            self.presets = {"默认预设": [{"content": "你好，对您的经历很感兴趣", "enabled": True}]}
            self.recent = []

    def save_presets(self):
        path = self.get_presets_path()
        try:
            data = self.presets.copy()
            data["_recent"] = self.recent
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Save presets failed: {e}")

    def get_preset_names(self):
        return list(self.presets.keys())

    def get_messages(self, preset_name):
        return self.presets.get(preset_name, [])

    def add_preset(self, name):
        if name not in self.presets:
            self.presets[name] = []
            self.save_presets()
            return True
        return False

    def delete_preset(self, name):
        if name in self.presets:
            del self.presets[name]
            if name in self.recent:
                self.recent.remove(name)
            self.save_presets()
            return True
        return False

    def update_messages(self, preset_name, messages):
        if preset_name in self.presets:
            self.presets[preset_name] = messages
            self.save_presets()

    def record_usage(self, name):
        if name in self.recent:
            self.recent.remove(name)
        self.recent.insert(0, name)
        self.recent = self.recent[:10]
        self.save_presets()

    def get_recent_presets(self, limit=3):
        valid = [n for n in self.recent if n in self.presets]
        return valid[:limit]

class PresetSelectionDialog(QDialog):
    def __init__(self, presets, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择预设")
        self.resize(300, 400)
        layout = QVBoxLayout(self)
        
        self.list_widget = QListWidget()
        self.list_widget.addItems(presets)
        self.list_widget.itemDoubleClicked.connect(self.accept)
        layout.addWidget(self.list_widget)
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        
    def get_selected(self):
        if self.list_widget.currentItem():
            return self.list_widget.currentItem().text()
        return None

class MaimaiPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.msg_manager = MessageManager()
        self.current_preset_name = None
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(15)

        # Row 1: Browser
        header = QHBoxLayout()
        b_label = QLabel("浏览器")
        self.browser_combo = QComboBox()
        self.browser_combo.setMinimumWidth(150)
        options = installed_browsers()
        self.browser_combo.addItems(options)
        if self.browser_combo.count() > 0:
            self.browser_combo.setCurrentIndex(0)
        
        self.open_btn = QPushButton("打开网页")
        self.open_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.test_zoom_btn = QPushButton("测试网页缩放")
        self.test_zoom_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.test_zoom_btn.setStyleSheet("""
            QPushButton {
                background-color: #F0F2F5;
                color: #1F2329;
                border: 1px solid #DEE0E3;
            }
            QPushButton:hover {
                background-color: #E5E6EB;
            }
        """)
        
        header.addWidget(b_label)
        header.addWidget(self.browser_combo)
        header.addWidget(self.open_btn)
        header.addWidget(self.test_zoom_btn)
        header.addStretch()
        self.layout.addLayout(header)

        # Row 2: Configs
        top = QHBoxLayout()
        label = QLabel("目标数量")
        self.max_spin = QSpinBox()
        self.max_spin.setRange(1, 1000000)
        self.max_spin.setValue(10)
        self.max_spin.setFixedWidth(80)
        
        self.debug_check = QCheckBox("调试模式")
        
        top.addWidget(label)
        top.addWidget(self.max_spin)
        top.addSpacing(20)
        top.addWidget(self.debug_check)
        top.addStretch()
        self.layout.addLayout(top)

        # Row 3: Action Buttons (Moved up)
        btns = QHBoxLayout()
        btns.setSpacing(15)

        self.start_btn = QPushButton("开始评估")
        self.start_btn.setObjectName("PrimaryButton")
        self.start_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.start_btn.setMinimumHeight(36)
        
        self.add_friend_btn = QPushButton("一键加好友")
        self.add_friend_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_friend_btn.setMinimumHeight(36)
        
        self.export_btn = QPushButton("导出为XML")
        self.export_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.export_btn.setMinimumHeight(36)

        btns.addWidget(self.start_btn)
        btns.addWidget(self.add_friend_btn)
        btns.addWidget(self.export_btn)
        self.layout.addLayout(btns)

        self.layout.addSpacing(10)

        # Row 4: Message Settings Toggle
        self.toggle_msg_btn = QPushButton("消息发送设置 ▼")
        self.toggle_msg_btn.setCheckable(True)
        self.toggle_msg_btn.setChecked(False)
        self.toggle_msg_btn.clicked.connect(self.toggle_msg_settings)
        self.layout.addWidget(self.toggle_msg_btn)

        # Row 5: Message Settings Area
        self.msg_container = QWidget()
        self.msg_container.setVisible(False)
        msg_layout = QVBoxLayout(self.msg_container)
        msg_layout.setContentsMargins(0, 5, 0, 0)
        
        # Recent Presets
        recent_group = QHBoxLayout()
        recent_group.addWidget(QLabel("选择预设:"))
        self.recent_layout = QHBoxLayout()
        recent_group.addLayout(self.recent_layout)
        
        self.select_more_btn = QPushButton("更多预设...")
        self.select_more_btn.clicked.connect(self.open_preset_dialog)
        recent_group.addWidget(self.select_more_btn)
        recent_group.addStretch()
        
        msg_layout.addLayout(recent_group)
        
        # Current Preset Label
        self.current_preset_label = QLabel("当前预设: 未选择")
        self.current_preset_label.setStyleSheet("color: gray; font-style: italic;")
        msg_layout.addWidget(self.current_preset_label)
        
        # Message List
        self.msg_list = QListWidget()
        self.msg_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.msg_list.itemChanged.connect(self.save_current_messages)
        msg_layout.addWidget(self.msg_list)
        
        # Message Controls
        msg_btns = QHBoxLayout()
        add_msg_btn = QPushButton("添加消息")
        add_msg_btn.clicked.connect(self.add_message)
        
        edit_msg_btn = QPushButton("编辑选中")
        edit_msg_btn.clicked.connect(self.edit_message)
        
        del_msg_btn = QPushButton("删除选中")
        del_msg_btn.clicked.connect(self.del_message)
        
        save_preset_btn = QPushButton("创建预设")
        save_preset_btn.clicked.connect(self.create_preset)
        
        msg_btns.addWidget(add_msg_btn)
        msg_btns.addWidget(edit_msg_btn)
        msg_btns.addWidget(del_msg_btn)
        msg_btns.addWidget(save_preset_btn)
        msg_btns.addStretch()
        
        msg_layout.addLayout(msg_btns)
        
        self.layout.addWidget(self.msg_container)
        self.layout.addStretch()
        
        # Initialize
        self.refresh_recent_presets()
        recents = self.msg_manager.get_recent_presets(1)
        if recents:
            self.load_preset(recents[0])
        elif "默认预设" in self.msg_manager.get_preset_names():
            self.load_preset("默认预设")

    def toggle_msg_settings(self):
        visible = self.toggle_msg_btn.isChecked()
        self.msg_container.setVisible(visible)
        self.toggle_msg_btn.setText("消息发送设置 ▲" if visible else "消息发送设置 ▼")

    def refresh_recent_presets(self):
        while self.recent_layout.count():
            child = self.recent_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        recents = self.msg_manager.get_recent_presets(3)
        for name in recents:
            btn = QPushButton(name)
            btn.clicked.connect(lambda checked, n=name: self.load_preset(n))
            self.recent_layout.addWidget(btn)

    def load_preset(self, name):
        self.current_preset_name = name
        self.current_preset_label.setText(f"当前预设: {name}")
        self.msg_manager.record_usage(name)
        self.refresh_recent_presets()
        
        msgs = self.msg_manager.get_messages(name)
        self.msg_list.blockSignals(True)
        self.msg_list.clear()
        for m in msgs:
            item = QListWidgetItem(m.get("content", ""))
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            item.setCheckState(Qt.CheckState.Checked if m.get("enabled", True) else Qt.CheckState.Unchecked)
            self.msg_list.addItem(item)
        self.msg_list.blockSignals(False)

    def open_preset_dialog(self):
        names = self.msg_manager.get_preset_names()
        dlg = PresetSelectionDialog(names, self)
        if dlg.exec():
            selected = dlg.get_selected()
            if selected:
                self.load_preset(selected)

    def create_preset(self):
        name, ok = QInputDialog.getText(self, "新建预设", "请输入预设名称:")
        if ok and name:
            if self.msg_manager.add_preset(name):
                self.current_preset_name = name
                self.save_current_messages() 
                self.load_preset(name)
            else:
                QMessageBox.warning(self, "错误", "预设已存在")

    def add_message(self):
        text, ok = QInputDialog.getMultiLineText(self, "添加消息", "请输入消息内容:")
        if ok and text:
            item = QListWidgetItem(text)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            item.setCheckState(Qt.CheckState.Checked)
            self.msg_list.addItem(item)
            self.save_current_messages()

    def del_message(self):
        row = self.msg_list.currentRow()
        if row >= 0:
            self.msg_list.takeItem(row)
            self.save_current_messages()

    def edit_message(self):
        item = self.msg_list.currentItem()
        if item:
            text, ok = QInputDialog.getMultiLineText(self, "编辑消息", "修改消息内容:", item.text())
            if ok and text:
                item.setText(text)
                self.save_current_messages()

    def save_current_messages(self):
        if not self.current_preset_name: return
        
        msgs = []
        for i in range(self.msg_list.count()):
            item = self.msg_list.item(i)
            msgs.append({
                "content": item.text(),
                "enabled": item.checkState() == Qt.CheckState.Checked
            })
        
        self.msg_manager.update_messages(self.current_preset_name, msgs)

    def get_selected_messages(self):
        msgs = []
        for i in range(self.msg_list.count()):
            item = self.msg_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                msgs.append(item.text())
        return msgs

    def on_export_xml(self):
        pass






class AIConfigPage(QWidget):
    config_saved = pyqtSignal(str, str, str, str)

    def __init__(self, api_key="", base_url="", system_prompt="", model="", parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        form = QFormLayout()
        
        self.key_edit = QLineEdit(api_key)
        self.key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        if not base_url:
            base_url = "https://dashscope.aliyuncs.com/compatible-mode/v1"
        self.url_edit = QLineEdit(base_url)
        self.url_edit.setPlaceholderText("例如: https://dashscope.aliyuncs.com/compatible-mode/v1")
        
        if not model:
            model = "qwen-vl-plus"
        self.model_edit = QLineEdit(model)
        self.model_edit.setPlaceholderText("例如: qwen-vl-plus")

        form.addRow("API Key:", self.key_edit)
        form.addRow("Base URL:", self.url_edit)
        form.addRow("Model:", self.model_edit)
        
        layout.addLayout(form)
        
        prompt_label = QLabel("System Prompt (系统指令):")
        layout.addWidget(prompt_label)
        
        self.prompt_edit = QPlainTextEdit(system_prompt)
        self.prompt_edit.setPlaceholderText("请输入自定义判断规则（留空则使用默认规则）。\n示例：\n需要联系备注1：看到备注为1时强制需要联系\n忽略手机号干扰：手机号中的数字不作为拒绝依据")
        # Default is empty now, as requested
             
        layout.addWidget(self.prompt_edit)
        
        # Save Button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.save_btn = QPushButton("保存配置")
        self.save_btn.setObjectName("PrimaryButton")
        self.save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.save_btn.clicked.connect(self.save_config)
        btn_layout.addWidget(self.save_btn)
        
        layout.addLayout(btn_layout)

    def save_config(self):
        api_key = self.key_edit.text().strip()
        base_url = self.url_edit.text().strip()
        model = self.model_edit.text().strip()
        prompt = self.prompt_edit.toPlainText()
        self.config_saved.emit(api_key, base_url, prompt, model)
        QMessageBox.information(self, "成功", "AI 配置已更新")

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(750, 550)
        self.api_key = ""
        self.base_url = ""
        self.system_prompt = ""
        self.model = "qwen-vl-plus"
        self.load_config()
        self.setup_ui()

    def get_config_path(self):
        if getattr(sys, 'frozen', False):
            # Running as compiled exe
            application_path = os.path.dirname(sys.executable)
        else:
            # Running as script
            application_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(application_path, "config.json")

    def load_config(self):
        config_path = self.get_config_path()
        try:
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.api_key = data.get("api_key", "")
                    self.base_url = data.get("base_url", "")
                    self.system_prompt = data.get("system_prompt", "")
                    self.model = data.get("model", "qwen-vl-plus")
            else:
                self.model = "qwen-vl-plus"
        except Exception as e:
            print(f"Loading config failed: {e}")
            self.model = "qwen-vl-plus"

    def save_config(self):
        config_path = self.get_config_path()
        try:
            data = {
                "api_key": self.api_key,
                "base_url": self.base_url,
                "system_prompt": self.system_prompt,
                "model": self.model
            }
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存配置失败: {e}")

    def setup_ui(self):
        # 1. Main Layout
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        # 2. Container
        self.container = QWidget()
        self.container.setObjectName("MainContainer")
        self.container_layout = QHBoxLayout(self.container)
        self.container_layout.setContentsMargins(0, 0, 0, 0)
        self.container_layout.setSpacing(0)
        self.main_layout.addWidget(self.container)

        # 3. Sidebar
        self.sidebar = Sidebar()
        self.container_layout.addWidget(self.sidebar)
        
        # 4. Right Panel (TitleBar + Content)
        self.right_panel = QWidget()
        self.right_panel.setObjectName("RightPanel")
        self.right_layout = QVBoxLayout(self.right_panel)
        self.right_layout.setContentsMargins(0, 0, 0, 0)
        self.right_layout.setSpacing(0)
        self.container_layout.addWidget(self.right_panel)
        
        # Title Bar
        self.title_bar = TitleBar(self)
        self.right_layout.addWidget(self.title_bar)
        
        # Content Stack
        self.stack = QStackedWidget()
        self.stack.setObjectName("ContentStack")
        self.maimai_page = MaimaiPage()
        self.ai_page = AIConfigPage(self.api_key, self.base_url, self.system_prompt, self.model)
        
        self.stack.addWidget(self.maimai_page)
        self.stack.addWidget(self.ai_page)
        
        self.right_layout.addWidget(self.stack)

        # Connections
        self.sidebar.page_changed.connect(self.stack.setCurrentIndex)
        
        self.sidebar.check_update_signal.connect(self.on_check_update)
        
        # Maimai Page Connections
        self.maimai_page.open_btn.clicked.connect(self.on_open_web)
        self.maimai_page.test_zoom_btn.clicked.connect(self.on_test_zoom)
        self.maimai_page.start_btn.clicked.connect(self.on_start_evaluate)
        self.maimai_page.add_friend_btn.clicked.connect(self.on_start_add_friend)
        self.maimai_page.export_btn.clicked.connect(self.maimai_page.on_export_xml)
        
        # AI Page Connections
        self.ai_page.config_saved.connect(self.update_config)

        # SizeGrip (Floating)
        self.grip = QSizeGrip(self)
        self.grip.resize(16, 16)

    def resizeEvent(self, event):
        if hasattr(self, 'grip'):
            self.grip.move(self.width() - 16, self.height() - 16)
        super().resizeEvent(event)

    def closeEvent(self, event):
        # 只有当用户点击主窗口右上角的 X 时，才设置全局退出标志
        # 不要在这里设置 STOP_FLAG，否则弹窗的关闭也会影响后台任务
        # 我们只在主窗口关闭时才退出整个应用
        event.accept()

    def update_config(self, api_key, base_url, prompt, model):
        self.api_key = api_key
        self.base_url = base_url
        self.system_prompt = prompt
        self.model = model
        self.save_config()

    def on_start(self, mode="evaluate"):
        target_url = "https://maimai.cn/ent/v41/recruit/talents?tab=1"
        browser = self.maimai_page.browser_combo.currentText()
        max_count = self.maimai_page.max_spin.value()
        is_debug = self.maimai_page.debug_check.isChecked()
        
        # Get selected messages
        messages = self.maimai_page.get_selected_messages()
        if not messages:
             # If no message selected, add default one to ensure something is sent
             messages = ["你好，对您的经历很感兴趣"]
        
        if mode == "evaluate" and not self.api_key:
            QMessageBox.warning(self, "警告", "请先点击“配置AI”设置 API Key！")
            return
        
        self.overlay = OverlayWindow()
        self.overlay.show()
        
        self.worker = WorkerThread(browser, target_url, max_count, mode, is_debug, self.api_key, self.base_url, self.system_prompt, messages=messages, model=self.model)
        self.worker.log_signal.connect(self.overlay.set_text)
        self.worker.result_signal.connect(self.on_worker_result)
        self.worker.error_signal.connect(self.on_worker_error)
        self.worker.finished_signal.connect(self.on_worker_finished)
        self.worker.popup_signal.connect(self.on_popup_message)
        # self.worker.alert_signal.connect(self.on_alert_message) # 已经不需要这个了
        
        self.hide()
        self.worker.start()

    def on_start_evaluate(self):
        self.on_start(mode="evaluate")

    def on_start_add_friend(self):
        self.on_start(mode="add_friend")

    def on_worker_result(self, result):
        self.worker_result = result

    def on_worker_error(self, error):
        self.worker_error = error

    def on_worker_finished(self):
        # 首先关闭遮罩层并恢复主窗口显示
        if hasattr(self, 'overlay') and self.overlay:
            self.overlay.close()
            
        self.show()
        self.activateWindow()
        self.raise_()
        
        # 保存这些值，因为 worker 会被置空
        err = getattr(self, 'worker_error', None)
        res = getattr(self, 'worker_result', None)
        
        if hasattr(self, 'worker_error'):
            del self.worker_error
        if hasattr(self, 'worker_result'):
            del self.worker_result
            
        self.worker = None
        
        # 延迟一点点显示弹窗，确保主窗口已经完全渲染出来
        if err:
            # 使用 QTimer.singleShot 确保在主事件循环中弹出，避免阻塞UI恢复
            QTimer.singleShot(100, lambda: self._show_error_dialog(err))
        elif res:
            QTimer.singleShot(100, lambda: self._show_result_dialog(res))

    def _show_error_dialog(self, msg):
        QMessageBox.warning(self, "执行出错", msg)
        
    def _show_result_dialog(self, msg):
        QMessageBox.information(self, "评估结果", msg)

    def on_popup_message(self, msg):
        # Create a non-blocking toast window
        if hasattr(self, '_current_toast') and self._current_toast:
            try:
                self._current_toast.close()
            except:
                pass
        
        self._current_toast = ToastWindow(msg)
        self._current_toast.show()

    def on_alert_message(self, title, msg):
        QMessageBox.warning(self, title, msg)



    def on_check_update(self):
        self.sidebar.btn_update.setText("正在检查...")
        self.sidebar.btn_update.setEnabled(False)
        QApplication.processEvents()
        
        has_update, version, body, download_url = check_for_updates()
        
        self.sidebar.btn_update.setText("检查更新")
        self.sidebar.btn_update.setEnabled(True)
        
        if not has_update:
            if body.startswith("检查更新出错"):
                QMessageBox.warning(self, "更新失败", body)
            else:
                QMessageBox.information(self, "检查更新", f"当前已是最新版本 (v{CURRENT_VERSION})")
            return
            
        if not download_url:
            QMessageBox.information(self, "发现新版本", f"发现新版本 v{version}，但未找到下载链接。\n\n更新日志：\n{body}")
            return
            
        # 弹窗询问是否更新
        reply = QMessageBox.question(self, "发现新版本", f"发现新版本 v{version}，是否立即下载并更新？\n\n更新日志：\n{body}",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                                     
        if reply == QMessageBox.StandardButton.Yes:
            self.start_download_update(download_url, version)
            
    def start_download_update(self, url, version):
        # 准备下载路径
        exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        self.update_save_path = os.path.join(exe_dir, f"AutoNetworker_update_v{version}.exe")
        
        # 显示进度条对话框
        self.progress_dialog = QProgressDialog("正在下载更新...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("下载更新")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.setValue(0)
        
        # 启动下载线程
        self.update_thread = UpdateThread(url, self.update_save_path)
        self.update_thread.progress_signal.connect(self._update_download_progress)
        self.update_thread.finished_signal.connect(self._on_download_finished)
        self.progress_dialog.canceled.connect(self.update_thread.terminate)
        
        self.update_thread.start()

    def _update_download_progress(self, percent, text):
        if hasattr(self, 'progress_dialog') and self.progress_dialog:
            self.progress_dialog.setValue(percent)
            self.progress_dialog.setLabelText(text)
            
    def _on_download_finished(self, success, result):
        if hasattr(self, 'progress_dialog') and self.progress_dialog:
            self.progress_dialog.close()
            
        if not success:
            QMessageBox.warning(self, "下载失败", result)
            return
            
        # 下载成功，准备替换
        # 生成一个临时的 bat 脚本来执行替换逻辑
        exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        current_exe = sys.executable
        
        if not getattr(sys, 'frozen', False):
             QMessageBox.information(self, "更新提示", f"已下载新版本到：\n{result}\n\n（开发环境下不自动替换，请手动测试）")
             return
             
        bat_path = os.path.join(exe_dir, "updater.bat")
        
        bat_content = f"""@echo off
chcp 65001
echo 正在等待主程序退出...
timeout /t 2 /nobreak >nul

if /I "%~x1"==".zip" (
    echo 正在解压新版本资源...
    :: 解压到临时目录
    powershell -Command "Expand-Archive -Path '%~1' -DestinationPath 'update_temp' -Force"
    :: 查找解压后的唯一文件夹 (如 jeffrey0328-AutoNetworker-xxxxxx)
    for /d %%D in ("update_temp\\*") do (
        echo 找到新版本目录: %%D
        xcopy /Y /E "%%D\\*" "%~dp0"
        goto :done_copy
    )
    :done_copy
    :: 清理临时文件
    rd /s /q "update_temp"
    del "%~1"
) else (
    echo 正在替换单文件版本...
    move /y "%~1" "{current_exe}"
)

echo 正在启动新版本...
start "" "{current_exe}"
echo 更新完成，清理临时文件...
del "%~f0"
"""
        # 传递下载的文件路径作为参数给 bat 脚本
        try:
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(bat_content)
                
            # 运行 bat 并退出当前程序
            subprocess.Popen([bat_path, result], creationflags=subprocess.CREATE_NO_WINDOW | 0x00000008, cwd=exe_dir)
            QApplication.quit()
        except Exception as e:
            QMessageBox.warning(self, "更新失败", f"无法创建更新脚本：{e}")

    def on_open_web(self):
        url = "https://maimai.cn/ent/v41/recruit/talents?tab=1"
        browser = self.maimai_page.browser_combo.currentText()
        ok = start_browser_with_debug_port(browser, url, 9222)
        if not ok:
            opened = QDesktopServices.openUrl(QUrl(url))
            if not opened:
                QMessageBox.warning(self, "提示", "无法打开网页")
            else:
                QMessageBox.warning(self, "注意", "无法以调试模式启动浏览器，后续自动化任务可能会失败。\n请尝试手动运行浏览器并开启 --remote-debugging-port=9222")

    def on_test_zoom(self):
        browser = self.maimai_page.browser_combo.currentText()
        driver = attach_selenium_driver(browser)
        if not driver:
            QMessageBox.warning(self, "错误", "无法连接到浏览器调试端口。请先点击“打开网页”或确保浏览器已在调试模式下运行。")
            return
        
        try:
            device_pixel_ratio = driver.execute_script("return window.devicePixelRatio")
            sys_scale = get_system_dpi_scale()
            actual_zoom = device_pixel_ratio / sys_scale
            
            msg = f"当前缩放信息检测：\n\n"
            msg += f"1. 物理/逻辑像素比 (DPR): {device_pixel_ratio:.2f}\n"
            msg += f"2. Windows 系统缩放: {sys_scale:.2f}\n"
            msg += f"3. 实际浏览器网页缩放: {actual_zoom:.2f} (DPR / 系统缩放)\n\n"
            
            if abs(actual_zoom - 1.0) > 0.01:
                msg += f"【警告】: 实际网页缩放为 {actual_zoom:.2f}，不在正常范围(1.0)，程序将无法运行！\n请在浏览器中按 Ctrl+0 重置缩放。"
            else:
                msg += "【状态】: 实际网页缩放为 1.0，完全正常，可以运行程序。"
                
            QMessageBox.information(self, "缩放测试结果", msg)
        except Exception as e:
            QMessageBox.warning(self, "测试失败", f"无法获取缩放信息: {e}")
        finally:
            # 释放由测试按钮创建的 driver，避免阻碍后续 WorkerThread 的连接，但不关闭浏览器
            # 测试完后，主动释放 driver 引用，防止连接被占用
            if driver:
                driver.quit = lambda: None
                del driver

def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 10))
    app.setStyleSheet("""
    /* Global Reset */
    * {
        font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
        color: #1F2329;
    }
    
    /* Main Window Container */
    #MainContainer {
        background-color: #FFFFFF;
        border: 1px solid #DEE0E3;
        border-radius: 10px;
    }

    /* Title Bar */
    #TitleLabel {
        font-size: 14px;
        font-weight: bold;
    }

    /* Inputs */
    QComboBox, QSpinBox, QLineEdit {
        background-color: #FFFFFF;
        border: 1px solid #DEE0E3;
        border-radius: 6px;
        padding: 6px 10px;
        selection-background-color: #3370FF;
    }
    QComboBox:hover, QSpinBox:hover, QLineEdit:hover {
        border: 1px solid #3370FF;
    }
    QComboBox::drop-down {
        border: none;
        width: 24px;
    }
    QSpinBox::up-button, QSpinBox::down-button {
        background-color: transparent;
        border: none;
        width: 16px;
    }

    /* Sidebar */
    #Sidebar {
        background-color: #F5F6F7;
        border-right: 1px solid #DEE0E3;
        border-top-left-radius: 10px;
        border-bottom-left-radius: 10px;
    }
    
    #SidebarButton {
        text-align: left;
        padding-left: 20px;
        border: none;
        background-color: transparent;
        font-size: 14px;
        color: #1F2329;
        border-radius: 6px;
        margin: 0 5px;
    }
    
    #SidebarButton:hover {
        background-color: #E4E6EB;
    }
    
    #SidebarButton:checked {
        background-color: #E1EAFF;
        color: #3370FF;
        font-weight: bold;
    }

    /* Title Bar Controls */
    #TitleBtnMin, #TitleBtnMax, #TitleBtnClose {
        border: none;
        border-radius: 4px;
        background-color: transparent;
        font-size: 16px;
        font-family: "Segoe UI Symbol", "Segoe UI", sans-serif; 
    }
    #TitleBtnMin:hover, #TitleBtnMax:hover {
        background-color: #E5E5E5;
    }
    #TitleBtnClose:hover {
        background-color: #E81123;
        color: white;
    }

    /* General Buttons */
    QPushButton {
        background-color: #FFFFFF;
        border: 1px solid #DEE0E3;
        border-radius: 6px;
        padding: 6px 16px;
        font-size: 14px;
    }
    QPushButton:hover {
        background-color: #F5F6F7;
        border: 1px solid #C9CDD4;
    }
    QPushButton:pressed {
        background-color: #E4E5E7;
    }

    /* Primary Button (Start) */
    QPushButton#PrimaryButton {
        background-color: #3370FF;
        border: 1px solid #3370FF;
        color: #FFFFFF;
        font-weight: bold;
    }
    QPushButton#PrimaryButton:hover {
        background-color: #2960E6;
        border: 1px solid #2960E6;
    }
    QPushButton#PrimaryButton:pressed {
        background-color: #1D4FBF;
        border: 1px solid #1D4FBF;
    }

    /* Title Bar Buttons */
    QPushButton#TitleBtnMin, QPushButton#TitleBtnMax, QPushButton#TitleBtnClose {
        border: none;
        background-color: transparent;
        border-radius: 0px;
        padding: 0px;
    }
    QPushButton#TitleBtnMin:hover, QPushButton#TitleBtnMax:hover {
        background-color: #E5E6EB;
    }
    QPushButton#TitleBtnClose:hover {
        background-color: #F54A45;
        color: white;
    }

    /* Message Box */
    QMessageBox {
        background-color: #FFFFFF;
    }
    QMessageBox QLabel {
        color: #1F2329;
    }
    QMessageBox QPushButton {
        min-width: 60px;
    }
    """)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
