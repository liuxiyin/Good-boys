import datetime
import os
import shutil
import time

import pandas as pd
from playwright.sync_api import sync_playwright


class OuZhi(object):
    """欧治"""

    # 初始化
    def __init__(self, playwright):
        # data目录创建
        self.downloads_path = os.path.join(os.getcwd(), 'data')  # 当前进程的工作目录，加上data，当做下载地址
        if not os.path.exists(self.downloads_path):  # 假如该目录不存在文件夹
            os.mkdir(self.downloads_path)  # 创建文件夹
        else:
            shutil.rmtree(self.downloads_path)  # 递归的删除文件夹
            os.mkdir(self.downloads_path)  # 创建文件夹

        # 初始化playwright
        self.playwright = playwright  # 将playwright变为类的self.playwright
        self.browser_type = self.playwright.chromium  # 创建浏览器对象，选择浏览器类型
        self.executable_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'   # 浏览器地址
        self.browser = self.browser_type.launch(headless=False, executable_path=self.executable_path,
                                                downloads_path=self.downloads_path)  # 创建同步对象，指定为有头模式，浏览器桌面地址与文件下载地址，自定义
        self.context = self.browser.new_context(accept_downloads=True)
        # 创建一个浏览器上下文，相当于一个独立的浏览器进程，不会与其他浏览器上下文共享cookie和缓存等信息
        self.page = self.context.new_page()  # 在浏览器上下文创建标签页

    # 登录
    def login(self):
        self.page.goto('http://202.109.133.33:8237/Index.html')  # 进入目标页面
        self.page.wait_for_load_state(state='networkidle')  # 等待登录页面的加载
        self.page.fill('//input[@id="account"]', setting["acc"])  # 填充账户名
        self.page.fill('//input[@id="pwd"]', setting["pwd"])  # 填充密码
        self.page.click('//button[@id="btnLogin"]')  # 点击登录
        self.page.wait_for_load_state(state='networkidle')  # 等待网页加载

    # 下载发货明细表
    def download(self):
        self.page.click('//a[text()="发货明细"]')  # 点击发货明细
        self.page.wait_for_load_state(state='networkidle')  # 等待网页加载

        start_date = (datetime.datetime.now() - datetime.timedelta(days=10)).strftime('%Y-%m-%d')  # 日期获取-十天前
        end_date = datetime.datetime.now().strftime('%Y-%m-%d')  # 日期获取-今日

        # 开始日期
        start_date_selector = self.page.query_selector('//input[@data-name="BeginDate"]')  # 获取符合条件的HTML元素
        self.page.evaluate(f'''start_date_selector => start_date_selector.value="{start_date}"''',
                           start_date_selector)  # 填充日期
        # 结束日期
        end_date_selector = self.page.query_selector('//input[@data-name="EndDate"]')  # 获取符合条件的HTML元素
        self.page.evaluate(f'''end_date_selector => end_date_selector.value="{end_date}"''',
                           end_date_selector)  # 填充日期

        self.page.click('//button[@id="btnSearch"]')  # 点击查询按钮
        self.page.wait_for_load_state(state='networkidle')
        if self.page.is_visible('//button[@i="close"]'):  # 如果存在该元素
            self.page.is_visible('//button[@i="close"]')
        time.sleep(1)  # 延迟一秒

        with self.page.expect_download() as download_info:
            self.page.click('//button[text()="生成报表"]')
        download = download_info.value
        file_name = download.suggested_filename
        self.down_load_file = os.path.join(self.downloads_path, file_name)  # 编写下载文件路径
        download.save_as(self.down_load_file)  # 下载文件保存刚编写的路径上

    # 主程序
    def main(self):
        self.login()
        self.download()


def screen_excel(excel_path):
    """
    欧治excel处理
    :param excel_path: excel路径
    :return:
    """
    # df['牌号'], df['规格']
    df = pd.read_excel(excel_path, keep_default_na=False)
    df = df[df['到站'] == '武钢港务外贸码头']  # 筛选到站为：武钢港务外贸码头
    # 切割货名
    df1 = df['货名'].str1.split(' ', expand=True)
    name = ['品名', '牌号', '规格', '_a']
    df1.columns = name
    df = df.join(df1)

    df.drop(labels=['货名', '_a'], axis=1, inplace=True)  # 删除货名列
    order = ['发货日期', '客户名称', '收货单位', '车船号', '件数', '理重', '磅重', '切边', '到站', '品名', '牌号', '规格', '降级', '等级', '长度']
    df = df[order]  # 排序
    df['规格'] = df['规格'].str1[1:]

    # 保存文件
    p, f = os.path.split(excel_path)
    result_path = os.path.join(p, 'result_' + f)
    df.to_excel(result_path, index=False)


if __name__ == '__main__':
    setting = {
        'acc': 'JG1172',
        'pwd': '838745'
    }
    with sync_playwright() as playwright:  # 打开同步对象
        ou_zhi = OuZhi(playwright)  # 把同步对象载入类，成为实例对象
        ou_zhi.main()  # 执行主程序
        excel_path = ou_zhi.down_load_file  # excel地址从实例对象中获取
        screen_excel(excel_path)
