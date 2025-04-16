#!/usr/bin/env python
# -*- coding: utf-8 -*-

import time
import random
import pickle
import csv
import os
from datetime import datetime, timedelta
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from wechatarticles import PublicAccountsWeb


# ============ 通用辅助函数 ============

def create_accounts_excel_file(filename="accounts.xlsx", example_accounts=None):
    """创建示例公众号名称Excel文件"""
    if example_accounts is None:
        example_accounts = ["同济计算机", "腾讯科技", "人民日报", "新华社", "CSDN"]
    
    # 如果文件已存在，不覆盖
    if os.path.exists(filename):
        print(f"文件 {filename} 已存在。")
        return
    
    # 创建DataFrame并写入Excel
    df = pd.DataFrame({"nickname": example_accounts})
    df.to_excel(filename, index=False)
    
    print(f"已创建示例公众号名称文件: {filename}")
    print(f"包含的公众号: {', '.join(example_accounts)}")


def read_accounts_from_excel(filename="accounts.xlsx"):
    """从Excel文件读取公众号名称列表"""
    if not os.path.exists(filename):
        print(f"文件 {filename} 不存在，将创建示例文件。")
        create_accounts_excel_file(filename)
    
    accounts = []
    try:
        # 读取Excel文件
        df = pd.read_excel(filename)
        
        # 检查是否有nickname列
        if 'nickname' in df.columns:
            # 过滤掉空值并转换为列表
            accounts = df['nickname'].dropna().astype(str).tolist()
        else:
            print(f"警告: Excel文件 {filename} 中没有找到'nickname'列")
    except Exception as e:
        print(f"读取 {filename} 时出错: {e}")
    
    print(f"从 {filename} 读取到 {len(accounts)} 个公众号")
    return accounts


def get_existing_article_titles(date=None):
    """获取指定日期Excel文件中的文章标题"""
    if date is None:
        # 使用昨天的日期
        date = datetime.now().date() - timedelta(days=1)
    
    # 根据日期生成文件名
    file_name = f"{date.month}月{date.day}号wechat_articles.xlsx"
    
    # 检查文件是否存在
    if not os.path.exists(file_name):
        print(f"昨天的文件 {file_name} 不存在，不需要检查重复文章")
        return set()
    
    try:
        # 读取Excel文件
        df = pd.read_excel(file_name, sheet_name='文章信息', skiprows=3)
        
        # 检查是否有title列
        if 'title' in df.columns:
            # 提取文章标题并转换为集合
            titles = set(df['title'].dropna().tolist())
            print(f"从昨天的文件 {file_name} 中读取到 {len(titles)} 个文章标题")
            return titles
        else:
            print(f"警告: Excel文件 {file_name} 中没有找到'title'列")
            return set()
    except Exception as e:
        print(f"读取昨天的文件 {file_name} 时出错: {e}")
        return set()


def save_articles_to_excel(articles_info, stats=None, output_file=None, filter_existing=True, stats_message=None):
    """将爬取的文章信息保存到Excel文件，可选择是否排除已存在的文章"""
    # 如果未指定输出文件名，则根据当前日期生成
    if output_file is None:
        current_date = datetime.now()
        output_file = f"{current_date.month}月{current_date.day}号wechat_articles.xlsx"
    
    filtered_articles = articles_info
    filtered_count = 0
    
    # 只有当需要过滤已存在文章时才执行
    if filter_existing:
        # 获取昨天Excel文件中的文章标题
        existing_titles = get_existing_article_titles()
        
        # 过滤掉已存在的文章
        filtered_articles = []
        
        for article in articles_info:
            if article['title'] in existing_titles:
                filtered_count += 1
                print(f"文章已存在于昨天的Excel中，将被过滤: {article['title']}")
            else:
                filtered_articles.append(article)
        
        print(f"过滤掉 {filtered_count} 篇已存在的文章，剩余 {len(filtered_articles)} 篇新文章")
    
    # 如果没有文章信息
    if not filtered_articles:
        print("没有文章信息可以保存")
        
        # 如果有统计信息，则创建一个Excel文件保存
        if stats:
            # 创建自定义统计信息
            if stats_message is None:
                stats_message = f"需要爬取的公众号一共{stats.get('total_accounts', 0)}个，"\
                              f"其中{stats.get('accounts_updated_recently', 0)}个最近有更新，"\
                              f"其中{stats.get('accounts_not_updated', 0)}个最近未更新，"\
                              f"过滤掉{filtered_count}篇已存在的文章"
                              
            df_stats = pd.DataFrame([{'统计信息': stats_message}])
            df_stats.to_excel(output_file, index=False)
            print(f"\n统计信息已保存到 {output_file}")
        return
    
    # 创建文章信息DataFrame
    df_articles = pd.DataFrame(filtered_articles)
    
    # 调整列的顺序，确保nickname, title, link, publish_time在前面
    preferred_columns = ['nickname', 'title', 'link', 'publish_time', 'publish_date']
    available_columns = [col for col in preferred_columns if col in df_articles.columns]
    other_columns = [col for col in df_articles.columns if col not in preferred_columns]
    all_columns = available_columns + other_columns
    df_articles = df_articles[all_columns]
    
    # 创建ExcelWriter对象
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 如果有统计信息，先写入统计信息
        if stats:
            # 创建自定义统计信息
            if stats_message is None:
                stats_message = f"需要爬取的公众号一共{stats.get('total_accounts', 0)}个，"\
                              f"其中{stats.get('accounts_updated_recently', 0)}个最近有更新，"\
                              f"其中{stats.get('accounts_not_updated', 0)}个最近未更新，"\
                              f"过滤掉{filtered_count}篇已存在的文章"
                              
            stats_row = pd.DataFrame([{'统计信息': stats_message}])
            stats_row.to_excel(writer, sheet_name='文章信息', index=False)
            
            # 写入文章信息，从第4行开始（给统计信息留出空间）
            df_articles.to_excel(writer, sheet_name='文章信息', startrow=3, index=False)
        else:
            # 没有统计信息，直接写入文章信息
            df_articles.to_excel(writer, sheet_name='文章信息', index=False)
    
    print(f"\n文章信息已保存到 {output_file}")
    if stats and stats_message:
        print(f"统计信息: {stats_message}")
    print(f"共保存了 {len(filtered_articles)} 篇文章")


# ============ 核心类：认证与凭证管理 ============

class WechatAuthManager:
    """微信公众平台凭证管理类"""
    
    def __init__(self, credentials_file="weixin_credentials.py"):
        """
        初始化认证管理器
        
        Args:
            credentials_file: 凭证文件路径
        """
        self.credentials_file = credentials_file
        self.cookie = None
        self.token = None
        self.crawler = None
    
    def load_credentials(self):
        """
        从文件加载凭证
        
        Returns:
            bool: 凭证是否有效
        """
        try:
            if not os.path.exists(self.credentials_file):
                print(f"凭证文件 {self.credentials_file} 不存在，需要登录获取")
                return False
            
            # 动态导入模块
            import importlib.util
            spec = importlib.util.spec_from_file_location("credentials", self.credentials_file)
            credentials = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(credentials)
            
            # 获取token和cookie
            self.token = getattr(credentials, 'token', '')
            self.cookie = getattr(credentials, 'cookie', '')
            
            if not self.token or not self.cookie:
                print("凭证文件中没有有效的token或cookie")
                return False
            
            print(f"已从 {self.credentials_file} 加载token和cookie")
            
            # 测试凭证是否有效
            return self.test_credentials()
        except Exception as e:
            print(f"加载凭证时出错: {e}")
            return False
    
    def test_credentials(self):
        """
        测试凭证是否有效
        
        Returns:
            bool: 凭证是否有效
        """
        print("正在测试保存的凭证是否有效...")
        
        if not self.cookie or not self.token:
            print("凭证不完整，无法测试")
            return False
        
        try:
            # 创建临时的PublicAccountsWeb实例
            web = PublicAccountsWeb(cookie=self.cookie, token=self.token)
            
            # 尝试获取一个公众号的信息（可以是任何存在的公众号）
            test_account = "微信公众平台"  # 这是一个官方账号，理论上一直存在
            result = web.get_urls(nickname=test_account, begin=0, count=1)
            
            # 检查返回结果
            if result is not None:
                print("保存的凭证有效，可以继续使用")
                return True
            else:
                print("保存的凭证无效，需要重新登录")
                return False
        except Exception as e:
            print(f"测试凭证时出错: {e}")
            print("将尝试重新登录获取新的凭证")
            return False
    
    def login_and_get_credentials(self, headless=False):
        """
        登录并获取凭证
        
        Args:
            headless: 是否使用无头模式
            
        Returns:
            bool: 登录是否成功
        """
        print("需要登录获取新的凭证...")
        self.crawler = WeixinMpCrawler(headless=headless)
        
        try:
            # 登录并获取凭证
            self.crawler.login()
            print("正在获取token和cookie...")
            time.sleep(5)  # 等待页面完全加载
            
            # 获取凭证
            self.token = self.crawler.get_token()
            self.cookie = self.crawler.get_cookie_string()
            
            if not self.token or not self.cookie:
                print("无法获取必要的凭证")
                return False
                
            # 打印凭证信息
            print("\n获取到的凭证:")
            print(f"token: {self.token}")
            print(f"cookie: {self.cookie[:50]}..." if len(self.cookie or '') > 50 else f"cookie: {self.cookie}")
            
            # 保存凭证到文件
            self.save_credentials()
            return True
        finally:
            if self.crawler:
                print("关闭登录浏览器...")
                self.crawler.close()
                
    def save_credentials(self):
        """将凭证保存到文件"""
        if not self.token or not self.cookie or not self.crawler:
            print("无法保存凭证：凭证不完整或爬虫未初始化")
            return False
            
        return self.crawler.save_credentials_to_py_file(self.credentials_file)
    
    def ensure_valid_credentials(self, headless=False):
        """
        确保有有效的凭证，必要时进行登录
        
        Args:
            headless: 是否使用无头模式（True表示不显示浏览器窗口）
            
        Returns:
            bool: 是否有有效凭证
        """
        # 尝试加载凭证
        credentials_valid = self.load_credentials()
        
        # 如果没有有效的凭证，进行登录
        if not credentials_valid:
            credentials_valid = self.login_and_get_credentials(headless=headless)
        
        return credentials_valid


# ============ 核心类：登录爬虫 ============

class WeixinMpCrawler:
    """微信公众平台登录爬虫，获取token和cookie"""
    
    def __init__(self, headless=False):
        """
        初始化爬虫
        
        Args:
            headless: 是否使用无头模式（不显示浏览器窗口）
        """
        self.url = "https://mp.weixin.qq.com/"
        self.browser = None
        self.headless = headless
        self.setup_browser()
        
    def setup_browser(self):
        """配置浏览器"""
        chrome_options = Options()
        if self.headless:
            chrome_options.add_argument("--headless")
        
        # 添加常用选项
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        
        # 初始化浏览器
        service = Service(ChromeDriverManager().install())
        self.browser = webdriver.Chrome(service=service, options=chrome_options)
        
    def login(self, timeout=120):
        """
        打开微信公众平台并等待用户扫码登录
        
        Args:
            timeout: 等待扫码登录的超时时间（秒）
            
        Returns:
            bool: 登录是否成功
        """
        try:
            print("正在打开微信公众平台登录页...")
            self.browser.get(self.url)
            
            # 等待二维码出现
            try:
                qrcode_iframe = WebDriverWait(self.browser, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "weui-desktop-qrcheck__iframe"))
                )
                self.browser.switch_to.frame(qrcode_iframe)
                
                # 等待二维码加载完成
                WebDriverWait(self.browser, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "qrcode"))
                )
                
                print("请使用微信扫描二维码登录（等待时间：{}秒）...".format(timeout))
                
                # 切回主frame
                self.browser.switch_to.default_content()
            except Exception as e:
                # 可能已经登录或页面结构变化
                print(f"获取二维码过程中出现异常: {str(e)}")
                print("可能已经登录或页面结构已变化，继续检测登录状态...")
            
            # 新的登录检测逻辑：不依赖于URL变化，而是检测页面上的元素
            start_time = time.time()
            while time.time() - start_time < timeout:
                # 检测是否已登录成功
                try:
                    # 尝试多种可能表示登录成功的元素
                    if (self.browser.find_elements(By.CLASS_NAME, "weui-desktop-panel__title") or
                        self.browser.find_elements(By.CLASS_NAME, "menu_item") or
                        self.browser.find_elements(By.ID, "menuBar") or
                        "/cgi-bin/home" in self.browser.current_url):
                        print("检测到登录成功标志！")
                        # 等待页面完全加载
                        time.sleep(3)
                        print("登录成功！")
                        return True
                except:
                    pass
                
                # 暂停一段时间再次检查
                time.sleep(2)
            
            # 如果超时但看起来已经登录（用户反馈），我们也认为成功
            print("登录检测超时，但将继续尝试获取token和cookie...")
            return True
            
        except Exception as e:
            print(f"登录过程中出现异常: {str(e)}")
            print("尽管出现错误，将继续尝试获取token和cookie...")
            # 即使出现异常，也尝试继续执行
            return True
    
    def get_cookies(self):
        """
        获取所有cookies
        
        Returns:
            list: cookie列表
        """
        return self.browser.get_cookies()
    
    def get_token(self):
        """
        尝试从页面或请求中提取token
        
        Returns:
            str: token字符串，如果未找到则返回None
        """
        try:
            print("尝试多种方法获取token...")
            
            # 方法1：从localStorage中获取
            try:
                tokens = self.browser.execute_script("""
                    var tokens = [];
                    for (var i = 0; i < localStorage.length; i++) {
                        var key = localStorage.key(i);
                        if (key.includes('token') || key.includes('Token')) {
                            tokens.push({
                                key: key,
                                value: localStorage.getItem(key)
                            });
                        }
                    }
                    return tokens;
                """)
                if tokens and len(tokens) > 0:
                    print(f"从localStorage找到可能的token: {tokens}")
                    # 返回第一个找到的token
                    return tokens[0]['value']
            except Exception as e:
                print(f"从localStorage获取token失败: {str(e)}")
                
            # 方法2：从URL中获取
            current_url = self.browser.current_url
            print(f"当前URL: {current_url}")
            if "token" in current_url.lower():
                # 使用简单解析从URL中提取token参数
                import re
                token_match = re.search(r'[?&](token|TOKEN)=([^&]+)', current_url, re.IGNORECASE)
                if token_match:
                    token = token_match.group(2)
                    print(f"从URL找到token: {token}")
                    return token
            
            # 方法3：从页面源码中查找
            page_source = self.browser.page_source
            if "token" in page_source.lower():
                # 使用更复杂的正则表达式查找不同格式的token
                import re
                patterns = [
                    r'"token":"([^"]+)"',
                    r'"Token":"([^"]+)"',
                    r'token:\s*["\']([^"\']+)["\']',
                    r'Token:\s*["\']([^"\']+)["\']',
                    r'token=([^&"\']+)',
                    r'Token=([^&"\']+)'
                ]
                
                for pattern in patterns:
                    token_match = re.search(pattern, page_source, re.IGNORECASE)
                    if token_match:
                        token = token_match.group(1)
                        print(f"从页面源码找到token: {token}")
                        return token
            
            # 方法4：尝试执行网络请求并从响应中提取
            try:
                token = self.browser.execute_script("""
                    // 创建一个简单的内部API请求来获取token
                    return new Promise((resolve, reject) => {
                        const xhr = new XMLHttpRequest();
                        xhr.open('GET', '/cgi-bin/bizattr?action=get_attr');
                        xhr.onload = function() {
                            if (this.status >= 200 && this.status < 300) {
                                try {
                                    const resp = JSON.parse(xhr.responseText);
                                    if (resp && resp.base_resp && resp.base_resp.token) {
                                        resolve(resp.base_resp.token);
                                    } else {
                                        resolve(null);
                                    }
                                } catch(e) {
                                    resolve(null);
                                }
                            } else {
                                resolve(null);
                            }
                        };
                        xhr.onerror = function() {
                            resolve(null);
                        };
                        xhr.send();
                    });
                """)
                if token:
                    print(f"从API响应找到token: {token}")
                    return token
            except Exception as e:
                print(f"从API获取token失败: {str(e)}")
            
            print("未找到token，可能需要登录后进入特定页面")
            return None
        except Exception as e:
            print(f"获取token时出错: {str(e)}")
            return None
    
    def get_cookie_string(self):
        """
        获取格式化的cookie字符串
        
        Returns:
            str: cookie字符串，格式为 'name1=value1; name2=value2'
        """
        cookies = self.get_cookies()
        return "; ".join([f"{cookie['name']}={cookie['value']}" for cookie in cookies])
    
    def save_cookies_to_file(self, filename="mp_weixin_cookies.pkl"):
        """
        将cookies保存到文件
        
        Args:
            filename: 保存的文件名
        """
        cookies = self.get_cookies()
        with open(filename, "wb") as f:
            pickle.dump(cookies, f)
        print(f"Cookies已保存到 {filename}")
    
    def save_credentials_to_py_file(self, filename="weixin_credentials.py"):
        """
        将token和cookie保存为Python变量格式的文件
        
        Args:
            filename: 保存的文件名
        """
        token = self.get_token()
        cookies = self.get_cookies()
        
        # 将cookie转换为请求可用的格式
        cookie_str = "; ".join([f"{cookie['name']}={cookie['value']}" for cookie in cookies])
        
        # 构建Python文件内容
        content = f"""#!/usr/bin/env python
# -*- coding: utf-8 -*-

# 微信公众平台凭证
# 自动生成于 {time.strftime("%Y-%m-%d %H:%M:%S")}

# 接口请求需要的token
token = "{token or ''}"

# 请求时需要携带的cookie字符串
cookie = '{cookie_str}'

# 可选：单独保存的cookie字典
cookie_dict = {{{', '.join([f'"{cookie["name"]}": "{cookie["value"]}"' for cookie in cookies])}}}
"""
        
        # 写入文件
        with open(filename, "w", encoding="utf-8") as f:
            f.write(content)
        
        print(f"\n凭证已保存到 {filename}")
        print(f"您可以通过 'import {filename.replace('.py', '')}' 导入并使用这些凭证")
        
        # 同时打印到控制台
        print("\n==== 可复制的凭证变量 ====")
        print(f"token = \"{token or ''}\"")
        print(f"cookie = '{cookie_str}'")
        
        return token, cookie_str
        
    def close(self):
        """关闭浏览器"""
        if self.browser:
            self.browser.quit()
            print("浏览器已关闭")


# ============ 核心类：文章爬取管理 ============

class ArticleCrawler:
    """微信公众号文章爬取管理类"""
    
    def __init__(self, cookie=None, token=None):
        """
        初始化文章爬取器
        
        Args:
            cookie: cookie字符串
            token: token字符串
        """
        self.cookie = cookie
        self.token = token
        self.web = None
        
        # 如果有cookie和token，就初始化web实例
        if self.cookie and self.token:
            self.init_web()
    
    def init_web(self):
        """初始化Web连接实例"""
        if self.cookie and self.token:
            self.web = PublicAccountsWeb(cookie=self.cookie, token=self.token)
            return True
        else:
            print("缺少必要的cookie或token，无法初始化连接")
            return False
    
    def set_credentials(self, cookie, token):
        """
        设置凭证
        
        Args:
            cookie: cookie字符串
            token: token字符串
        """
        self.cookie = cookie
        self.token = token
        return self.init_web()
    
    def extract_publish_time_from_url(self, url):
        """
        从微信公众号文章URL中提取发布时间
        
        Args:
            url: 文章URL
            
        Returns:
            tuple: (发布日期对象, 发布日期字符串)
        """
        try:
            # 设置请求头，模拟浏览器访问
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
                'Connection': 'keep-alive'
            }
            
            # 随机延迟，避免请求过于频繁
            time.sleep(random.uniform(1, 3))
            
            # 请求文章页面获取内容
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                # 使用BeautifulSoup解析HTML
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # 1. 尝试查找微信文章页面中的发布时间元素
                publish_time_element = soup.select_one('#publish_time') or soup.select_one('.publish_time')
                if publish_time_element:
                    publish_time_text = publish_time_element.text.strip()
                    # 尝试解析日期
                    date_formats = ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%Y年%m月%d日 %H:%M', '%Y年%m月%d日']
                    for date_format in date_formats:
                        try:
                            publish_date = datetime.strptime(publish_time_text, date_format).date()
                            return publish_date, publish_time_text
                        except ValueError:
                            continue
                
                # 2. 在页面中查找时间模式
                date_patterns = [
                    r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}',
                    r'\d{4}-\d{2}-\d{2}',
                    r'\d{4}年\d{1,2}月\d{1,2}日 \d{1,2}:\d{1,2}',
                    r'\d{4}年\d{1,2}月\d{1,2}日'
                ]
                
                # 遍历页面查找可能的日期
                page_text = response.text
                for pattern in date_patterns:
                    matches = re.findall(pattern, page_text)
                    if matches:
                        # 尝试解析找到的第一个日期
                        for match in matches:
                            date_formats = ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%Y年%m月%d日 %H:%M', '%Y年%m月%d日']
                            for date_format in date_formats:
                                try:
                                    publish_date = datetime.strptime(match, date_format)
                                    return publish_date.date(), match
                                except ValueError:
                                    continue
            
            # 如果从页面内容无法获取到时间，返回None
            print(f"无法从URL {url} 提取发布时间")
            return None, ""
            
        except Exception as e:
            print(f"提取文章发布时间时出错: {e}")
            return None, ""
    
    def fetch_articles_from_account(self, nickname, count=10, filter_recent_days=None, max_attempts=10, time_filter_func=None, stop_on_outdated=False):
        """
        从单个公众号获取文章
        
        Args:
            nickname: 公众号名称
            count: 获取的文章数量
            filter_recent_days: 如果不为None，只获取最近几天的文章
            max_attempts: 最大尝试次数
            time_filter_func: 时间过滤函数，接收article_date参数，返回布尔值
            stop_on_outdated: 是否在发现第一篇过期文章时就停止，适用于批量爬取多个公众号时
            
        Returns:
            list: 文章信息列表 [{nickname, title, link, publish_time}, ...]
        """
        if not self.web:
            print("Web连接未初始化，请先设置有效凭证")
            return []
            
        articles_info = []
        
        # 设置日期范围（如果需要）
        today = datetime.now().date()
        date_range = None
        if filter_recent_days:
            date_range = [today - timedelta(days=i) for i in range(filter_recent_days)]
            print(f"将只保留最近 {filter_recent_days} 天的文章")
        
        # 使用随机延迟，避免操作过于规律被检测
        delay = random.uniform(3, 8)
        print(f"等待 {delay:.2f} 秒后开始获取文章...")
        time.sleep(delay)
        
        attempt = 0
        offset = 0
        collected = 0
        has_more = True
        outdated_found = False  # 标记是否找到过期文章
        
        # 最大批次，每批次获取数量
        batch_size = 5  # 减小批次大小，减少漏爬风险
        
        # 记录连续空结果次数
        empty_results_count = 0
        
        # 记录已获取的文章链接，避免重复
        fetched_links = set()
        
        while has_more and collected < count and attempt < max_attempts and not outdated_found:
            try:
                print(f"获取公众号 '{nickname}' 的文章，批次 {attempt+1}，偏移量 {offset}...")
                
                # 获取文章数据
                articles = self.web.get_urls(nickname=nickname, begin=offset, count=batch_size)
                
                if not articles:
                    empty_results_count += 1
                    print(f"未获取到文章，连续空结果次数: {empty_results_count}")
                    
                    # 只有连续3次获取不到文章，才认为确实没有更多文章了
                    if empty_results_count >= 3:
                        print(f"连续{empty_results_count}次未获取到文章，可能已到达文章列表末尾")
                        has_more = False
                        break
                    else:
                        # 尽管没有获取到文章，仍然尝试增加偏移量继续获取
                        # 使用较小的增量，避免跳过文章
                        print("尝试增加偏移量继续获取...")
                        offset += 1  # 只增加1而不是batch_size，更小的增量减少漏爬
                        attempt += 1
                        
                        # 添加额外延迟，可能是因为请求过于频繁导致的限制
                        extra_delay = random.uniform(8, 15)
                        print(f"添加额外延迟 {extra_delay:.2f} 秒...")
                        time.sleep(extra_delay)
                        continue
                else:
                    # 重置连续空结果计数
                    empty_results_count = 0
                    
                print(f"获取到 {len(articles)} 篇文章，正在处理...")
                
                # 标记是否在此批次中添加了任何文章
                added_in_batch = False
                
                # 处理获取到的文章
                for article in articles:
                    # 从文章数据中提取所需信息
                    title = article.get('title', '无标题')
                    link = article.get('link', '无链接')
                    
                    # 检查是否已经处理过这个链接
                    if link in fetched_links:
                        print(f"文章 '{title}' 已经爬取过，跳过")
                        continue
                    
                    # 从URL中提取发布时间
                    article_date = None
                    publish_date_str = ''
                    
                    if link != '无链接':
                        article_date, publish_date_str = self.extract_publish_time_from_url(link)
                        
                        if article_date:
                            # 如果有自定义的时间过滤函数
                            if time_filter_func and not time_filter_func(article_date):
                                print(f"文章 '{title}' 不符合时间过滤条件，跳过")
                                
                                # 如果设置了stop_on_outdated，并且文章确实是因为太旧而被过滤
                                # 这里我们假设time_filter_func是用来检查文章是否在最近的日期范围内
                                if stop_on_outdated:
                                    print(f"发现不符合时间条件的文章，停止爬取公众号 '{nickname}'")
                                    outdated_found = True
                                    break
                                    
                                continue
                                
                            # 如果需要过滤最近几天的文章
                            if date_range and article_date not in date_range:
                                print(f"文章 '{title}' 发布于 {article_date}，不在指定的日期范围内")
                                
                                # 无论是否已收集文章，只要文章日期早于范围，如果启用了stop_on_outdated，就停止
                                if stop_on_outdated:
                                    print(f"发现超出日期范围的文章，停止爬取公众号 '{nickname}'")
                                    outdated_found = True
                                    break
                                    
                                # 原有的逻辑：只有已经收集了文章才因为日期早于范围而停止
                                if article_date < min(date_range) and collected > 0:
                                    print("已找到早于指定日期范围的文章，停止获取")
                                    has_more = False
                                    break
                                continue
                            
                            # 添加文章信息
                            articles_info.append({
                                'nickname': nickname,
                                'title': title,
                                'link': link,
                                'publish_time': publish_date_str,
                                'publish_date': article_date.strftime('%Y-%m-%d')
                            })
                            
                            # 记录已获取的链接
                            fetched_links.add(link)
                            
                            collected += 1
                            added_in_batch = True
                            print(f"[{collected}/{count}] 已添加文章: {title}")
                            
                            # 如果已经收集足够的文章，则终止循环
                            if collected >= count:
                                print(f"已达到目标数量 {count} 篇文章，停止获取")
                                has_more = False
                                break
                        else:
                            print(f"从URL无法提取到发布时间: {link}")
                
                # 如果发现了过期文章，退出主循环
                if outdated_found:
                    print(f"由于发现过期文章，提前停止爬取公众号 '{nickname}'")
                    break
                    
                # 更新偏移量，准备获取下一批文章
                if added_in_batch:
                    # 如果这批次成功添加了文章，使用常规偏移量增量
                    offset += batch_size // 2  # 使用一半的批次大小作为增量，提高重叠度
                else:
                    # 如果这批次没有添加任何文章，使用较小的增量尝试
                    offset += 1  # 增加最小偏移量，尝试捕获可能漏掉的文章
                
                # 如果还有更多文章需要获取，添加随机延迟
                if has_more and collected < count:
                    delay = random.uniform(5, 10)
                    print(f"等待 {delay:.2f} 秒后获取下一批文章...")
                    time.sleep(delay)
                    
                attempt += 1
                    
            except Exception as e:
                attempt += 1
                print(f"获取公众号 '{nickname}' 的文章时出错 (尝试 {attempt}/{max_attempts}): {e}")
                if attempt < max_attempts:
                    delay = random.uniform(10, 15)
                    print(f"等待 {delay:.2f} 秒后重试...")
                    time.sleep(delay)
                else:
                    print(f"已达到最大尝试次数 {max_attempts}，停止获取")
                    break
        
        if outdated_found:
            print(f"提前终止：公众号 '{nickname}' 的文章已超出时间范围，共获取到 {len(articles_info)} 篇符合条件的文章")
        else:
            print(f"共获取到公众号 '{nickname}' 的 {len(articles_info)} 篇文章")
            
        return articles_info
    
    def fetch_wechat_articles(self, nickname_list, articles_per_account=10, days=2):
        """
        爬取多个公众号的文章
        
        Args:
            nickname_list: 公众号名称列表
            articles_per_account: 每个公众号获取的文章数量(最大值)
            days: 获取最近几天的文章 (默认为2，即今天和昨天)
            
        Returns:
            tuple: (文章信息列表 [{nickname, title, link, publish_time}, ...], 统计信息)
        """
        if not self.web:
            print("Web连接未初始化，请先设置有效凭证")
            return [], {}
            
        # 获取当前日期和昨天日期
        today = datetime.now().date()
        date_range = [today - timedelta(days=i) for i in range(days)]
        
        print(f"当前日期: {today}，将抓取最近 {days} 天发布的文章")
        
        # 存储所有文章信息的列表
        all_articles_info = []
        
        # 统计信息
        total_accounts = len(nickname_list)
        accounts_updated_recently = 0
        accounts_not_updated = 0
        
        # 定义时间过滤函数：只保留最近days天的文章
        def recent_days_filter(article_date):
            return article_date in date_range
        
        for i, nickname in enumerate(nickname_list):
            print(f"\n正在获取公众号 '{nickname}' 的文章 ({i+1}/{total_accounts})...")
            
            # 获取该公众号的文章（使用时间过滤）
            # 启用stop_on_outdated，一旦发现过期文章就停止爬取当前公众号
            account_articles = self.fetch_articles_from_account(
                nickname=nickname,
                count=articles_per_account,
                time_filter_func=recent_days_filter,
                stop_on_outdated=True  # 添加这个参数，一旦发现过期文章就停止
            )
            
            # 添加文章到总列表
            all_articles_info.extend(account_articles)
            
            # 更新统计信息
            if account_articles:
                accounts_updated_recently += 1
                print(f"公众号 '{nickname}' 最近 {days} 天有更新，找到 {len(account_articles)} 篇文章")
            else:
                accounts_not_updated += 1
                print(f"公众号 '{nickname}' 最近 {days} 天无更新")
            
            # 在每个公众号处理后添加额外的随机延迟
            if i < len(nickname_list) - 1:  # 如果不是最后一个公众号
                extra_delay = random.uniform(8, 15)
                print(f"处理下一个公众号前等待 {extra_delay:.2f} 秒...")
                time.sleep(extra_delay)
        
        # 准备统计信息
        stats = {
            'total_accounts': total_accounts,
            'accounts_updated_recently': accounts_updated_recently,
            'accounts_not_updated': accounts_not_updated,
            'date': today.strftime('%Y-%m-%d')
        }
        
        return all_articles_info, stats
    
    def fetch_account_history(self, nickname, max_articles=100):
        """
        爬取单个公众号的历史文章（最多max_articles篇）
        
        Args:
            nickname: 公众号名称
            max_articles: 最大爬取文章数量，默认100
            
        Returns:
            list: 文章信息列表 [{nickname, title, link, publish_time, publish_date}, ...]
        """
        if not self.web:
            print("Web连接未初始化，请先设置有效凭证")
            return []
            
        print(f"===== 开始爬取公众号 '{nickname}' 的历史文章 =====")
        
        # 获取该公众号的文章（不使用时间过滤）
        articles = self.fetch_articles_from_account(
            nickname=nickname,
            count=max_articles
        )
        
        print(f"===== 完成爬取公众号 '{nickname}' 的历史文章，共获取 {len(articles)} 篇 =====")
        return articles


# ============ 核心类：文章内容分析 ============

class ArticleAnalyzer:
    """文章内容分析类"""
    
    def __init__(self):
        """初始化分析器"""
        pass
        
    def fetch_article_content(self, url):
        """
        从文章链接获取完整内容
        
        Args:
            url: 文章URL
            
        Returns:
            str: 文章内容文本
        """
        try:
            # 设置请求头，模拟浏览器访问
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
                'Connection': 'keep-alive'
            }
            
            # 随机延迟，避免请求过于频繁
            time.sleep(random.uniform(1, 3))
            
            # 请求文章页面获取内容
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                # 使用BeautifulSoup解析HTML
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # 提取文章内容 - 微信文章通常在rich_media_content中
                content_element = soup.select_one('#js_content') or soup.select_one('.rich_media_content')
                if content_element:
                    # 获取所有文本，清理空白字符
                    content_text = content_element.get_text(strip=True)
                    return content_text
                
                # 如果找不到特定元素，尝试获取整个页面文本
                return soup.get_text(strip=True)
            
            # 如果请求失败
            print(f"请求文章内容失败，状态码: {response.status_code}")
            return ""
            
        except Exception as e:
            print(f"获取文章内容时出错: {e}")
            return ""
    
    def calculate_keyword_score(self, text, keywords, weights):
        """
        计算文本中关键词得分
        
        Args:
            text: 文本内容
            keywords: 关键词列表
            weights: 权重列表，与关键词一一对应
            
        Returns:
            tuple: (关键词计数字典, 总分数)
        """
        # 确保关键词和权重长度一致
        if len(keywords) != len(weights):
            print("关键词和权重数量不匹配")
            return {}, 0
        
        # 统计每个关键词出现次数
        keyword_counts = {}
        total_score = 0
        
        for i, keyword in enumerate(keywords):
            # 统计关键词出现次数 (不区分大小写)
            count = text.lower().count(keyword.lower())
            keyword_counts[keyword] = count
            
            # 计算加权分数
            score = count * weights[i]
            total_score += score
        
        return keyword_counts, total_score
    
    def analyze_articles_with_keywords(self, articles, keywords, weights):
        """
        分析文章列表中的关键词
        
        Args:
            articles: 文章信息列表
            keywords: 关键词列表
            weights: 权重列表
            
        Returns:
            list: 分析后的文章列表 (添加了keyword_counts和keyword_score字段)
        """
        # 限制关键词数量
        if len(keywords) > 3:
            print("关键词数量超过限制，只使用前3个关键词")
            keywords = keywords[:3]
            weights = weights[:3]
        
        # 如果权重列表不够长，用1补齐
        while len(weights) < len(keywords):
            weights.append(1)
            
        print(f"开始分析 {len(articles)} 篇文章中的关键词: {keywords}")
        print(f"关键词权重: {weights}")
        
        # 为每篇文章获取内容并计算关键词分数
        for i, article in enumerate(articles):
            print(f"[{i+1}/{len(articles)}] 处理文章: {article['title']}")
            
            # 获取文章内容
            content = self.fetch_article_content(article['link'])
            article['content_length'] = len(content)
            
            # 计算关键词分数
            keyword_counts, total_score = self.calculate_keyword_score(content, keywords, weights)
            
            # 保存到文章信息
            article['keyword_counts'] = str(keyword_counts)  # 转为字符串以便保存到Excel
            article['keyword_score'] = total_score
            
            # 打印分数
            print(f"- 关键词统计: {keyword_counts}")
            print(f"- 总分数: {total_score}")
            
            # 添加随机延迟，避免请求过于频繁
            if i < len(articles) - 1:  # 如果不是最后一篇文章
                delay = random.uniform(2, 5)
                print(f"等待 {delay:.2f} 秒后处理下一篇文章...")
                time.sleep(delay)
        
        # 根据关键词得分排序文章
        sorted_articles = sorted(articles, key=lambda x: x.get('keyword_score', 0), reverse=True)
        
        print(f"===== 完成关键词分析，按分数排序 =====")
        for i, article in enumerate(sorted_articles[:10]):  # 打印前10篇
            if i < len(sorted_articles):
                print(f"{i+1}. {article['title']} - 分数: {article.get('keyword_score', 0)}")
        
        return sorted_articles


# ============ 高级封装类：微信文章管理器 ============

class WechatArticleManager:
    """微信公众号文章管理器 - 高级封装类"""
    
    def __init__(self, credentials_file="weixin_credentials.py", headless=False):
        """
        初始化管理器
        
        Args:
            credentials_file: 凭证文件路径
            headless: 是否使用无头模式（True表示不显示浏览器窗口）
        """
        self.auth_manager = WechatAuthManager(credentials_file)
        self.crawler = None
        self.analyzer = ArticleAnalyzer()
        self.headless = headless  # 保存无头模式设置
    
    def ensure_authentication(self):
        """
        确保身份验证有效
        
        Returns:
            bool: 身份验证是否成功
        """
        if self.auth_manager.ensure_valid_credentials(headless=self.headless):
            # 创建或更新爬虫实例
            if not self.crawler:
                self.crawler = ArticleCrawler(self.auth_manager.cookie, self.auth_manager.token)
            else:
                self.crawler.set_credentials(self.auth_manager.cookie, self.auth_manager.token)
            return True
        else:
            print("无法获取有效凭证，操作中止")
            return False
    
    def crawl_multiple_accounts(self, nickname_list, articles_per_account=10, days=2, output_file=None):
        """
        爬取多个公众号的最近文章
        
        Args:
            nickname_list: 公众号名称列表
            articles_per_account: 每个公众号获取的文章数量
            days: 获取最近几天的文章
            output_file: 输出文件名，默认为None(自动生成)
            
        Returns:
            tuple: (成功标志, 文章列表)
        """
        # 确保认证有效
        if not self.ensure_authentication():
            return False, []
        
        # 爬取文章
        articles, stats = self.crawler.fetch_wechat_articles(
            nickname_list, 
            articles_per_account=articles_per_account, 
            days=days
        )
        
        if not articles:
            print("未获取到任何文章")
            return False, []
        
        # 保存到Excel
        if output_file is None:
            current_date = datetime.now()
            output_file = f"{current_date.month}月{current_date.day}号wechat_articles.xlsx"
            
        save_articles_to_excel(articles, stats, output_file)
        
        print(f"\n爬取完成！共爬取了 {len(articles)} 篇最近 {days} 天发布的文章")
        print(f"数据已保存到 {output_file}")
        
        return True, articles
    
    def crawl_account_history(self, nickname, max_articles=100, output_file=None):
        """
        爬取单个公众号的历史文章
        
        Args:
            nickname: 公众号名称
            max_articles: 最大爬取文章数量
            output_file: 输出文件名，默认为None(自动生成)
            
        Returns:
            tuple: (成功标志, 文章列表)
        """
        # 确保认证有效
        if not self.ensure_authentication():
            return False, []
        
        # 爬取文章
        articles = self.crawler.fetch_account_history(nickname, max_articles)
        
        if not articles:
            print(f"未获取到公众号 '{nickname}' 的任何文章")
            return False, []
        
        # 如果未指定输出文件名，则使用公众号名称自动生成
        if output_file is None:
            current_date = datetime.now()
            output_file = f"{nickname}_{current_date.strftime('%Y%m%d')}_历史文章.xlsx"
        
        # 创建简单的统计信息
        stats_message = f"公众号 '{nickname}' 历史文章爬取结果，共获取 {len(articles)} 篇文章"
        
        # 保存到Excel（不过滤已存在的文章）
        save_articles_to_excel(
            articles_info=articles,
            output_file=output_file,
            filter_existing=False,
            stats_message=stats_message
        )
        
        print(f"\n爬取完成！共爬取了公众号 '{nickname}' 的 {len(articles)} 篇历史文章")
        print(f"数据已保存到 {output_file}")
        
        return True, articles
    
    def search_keywords_in_account(self, nickname, keywords, weights=None, max_articles=20, output_file=None):
        """
        搜索关键词并排序公众号文章
        
        Args:
            nickname: 公众号名称
            keywords: 关键词列表
            weights: 权重列表，默认都为1
            max_articles: 最大爬取文章数量
            output_file: 输出文件名，默认为None(自动生成)
            
        Returns:
            tuple: (成功标志, 排序后的文章列表)
        """
        # 默认权重
        if weights is None:
            weights = [1] * len(keywords)
            
        # 确保认证有效
        if not self.ensure_authentication():
            return False, []
        
        # 爬取文章
        articles = self.crawler.fetch_account_history(nickname, max_articles)
        
        if not articles:
            print(f"未获取到公众号 '{nickname}' 的任何文章")
            return False, []
        
        # 分析关键词并排序
        sorted_articles = self.analyzer.analyze_articles_with_keywords(articles, keywords, weights)
        
        # 如果未指定输出文件名，则使用关键词和公众号名称自动生成
        if output_file is None:
            keyword_str = "_".join(keywords)
            current_date = datetime.now()
            output_file = f"{nickname}_{keyword_str}_{current_date.strftime('%Y%m%d')}.xlsx"
        
        # 创建简单的统计信息
        keywords_str = ", ".join([f"{keywords[i]}(权重{weights[i]})" for i in range(len(keywords))])
        stats_message = f"公众号 '{nickname}' 关键词搜索: {keywords_str}，共分析 {len(sorted_articles)} 篇文章"
        
        # 保存到Excel（不过滤已存在的文章）
        save_articles_to_excel(
            articles_info=sorted_articles,
            output_file=output_file,
            filter_existing=False,
            stats_message=stats_message
        )
        
        print(f"\n搜索完成！共处理公众号 '{nickname}' 的 {len(sorted_articles)} 篇文章")
        print(f"数据已按关键词分数从高到低排序并保存到 {output_file}")
        
        return True, sorted_articles


# ============ 如果直接运行此脚本 ============

if __name__ == "__main__":
    # 程序入口示例
    manager = WechatArticleManager()
    
    # 示例：搜索关键词
    target_account = "同济计算机"  # 修改为您想爬取的公众号名称
    keywords = ["人工智能", "数据科学", "程序设计"]
    weights = [1.5, 1.2, 1.0]
    
    success, articles = manager.search_keywords_in_account(
        nickname=target_account,
        keywords=keywords,
        weights=weights,
        max_articles=20
    )
