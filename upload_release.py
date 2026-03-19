import os
import sys
import json
import requests
import logging
from datetime import datetime

# 设置日志配置
LOG_FILE = "upload_log.txt"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# 配置信息 (可以直接从 main.py 读取或在此处写死)
_TOKEN_PART_1 = "github_pat_11AM5HAJY0D1cABcJKoSyk_nJ4XHE"
_TOKEN_PART_2 = "BJZ7AmMpsNMwSAwVos9A4Sun6c8qzIzcYQZwvBIOHALEMF4qiq679"
GITHUB_TOKEN = _TOKEN_PART_1 + _TOKEN_PART_2
REPO_OWNER = "jeffrey0328"
REPO_NAME = "AutoNetworker"

def get_current_version():
    """从 main.py 中提取 CURRENT_VERSION"""
    try:
        with open("main.py", "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("CURRENT_VERSION"):
                    # CURRENT_VERSION = "1.0.0" -> 1.0.0
                    return line.split("=")[1].strip().strip('"').strip("'")
    except Exception as e:
        logging.error(f"读取版本号失败: {e}")
    return None

def create_and_upload_release():
    logging.info("====================================")
    logging.info("开始执行发布流程")
    
    version = get_current_version()
    if not version:
        logging.error("无法获取版本号，上传取消。")
        return

    tag_name = f"v{version}"
    logging.info(f"检测到版本号: {version} -> Tag: {tag_name}")

    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28"
    }

    # 1. 检查 tag 是否已经存在 (防止重复创建报错)
    releases_url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases"
    logging.info(f"正在请求 GitHub API 创建 Release: {releases_url}")
    
    release_data = {
        "tag_name": tag_name,
        "name": f"Release {tag_name}",
        "body": f"自动打包发布的新版本 v{version}\n\n发布时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "draft": False,
        "prerelease": False
    }

    try:
        resp = requests.post(releases_url, headers=headers, json=release_data)
        
        if resp.status_code == 201:
            release_info = resp.json()
            html_url = release_info["html_url"]
            logging.info(f"Release {tag_name} 创建成功！")
            logging.info(f"发布页面地址: {html_url}")
            logging.info("提示: 程序源文件已直接推送到 main 分支，用户将通过 zipball_url 自动更新。无需手动上传附件！")
            
            # 尝试自动打开浏览器 (可选)
            try:
                os.startfile(html_url)
            except Exception:
                pass
                
        elif resp.status_code == 422 and "already exists" in resp.text:
            logging.warning(f"版本 {tag_name} 已经存在！")
            return
        else:
            logging.error(f"创建 Release 失败: {resp.status_code} - {resp.text}")
            logging.error("提示: 403 错误通常是因为 Token 权限不足。请确保 Token 具有 'Contents: Read and Write' 权限。")
            return

        # 2. 移除自动上传逻辑
        # logging.info(f"准备上传文件: {FILE_TO_UPLOAD}")
        # ... (后续上传代码已移除)

    except requests.exceptions.Timeout:
        logging.error("请求超时！")
    except requests.exceptions.RequestException as e:
        logging.error(f"网络错误: {e}")
    except Exception as e:
        logging.exception(f"发生未知错误: {e}")

if __name__ == "__main__":
    create_and_upload_release()
    os.system("pause")