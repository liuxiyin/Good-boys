"""
蝉妈妈直播数据拉取
"""
import os

from playwright.sync_api import sync_playwright
from loguru import logger
import pandas as pd
import time
import uuid
import random
# import datetime
import email.mime.multipart
import email.mime.text
import smtplib
from email.mime.application import MIMEApplication


# ex_path = "./ms-playwright/chromium-907428/chrome-win/chrome"
ex_path = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
ex_path = ex_path if os.path.exists(ex_path) else None


setting = {
    'account': '13524856993',
    'pwd': '123456a',
    'url': 'https://www.chanmama.com/login'

}


class ChanMaMa:
    def __init__(self):
        self.base_url = "https://www.chanmama.com/report/detail/"
        self.data = []
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) + '.xlsx'  # uuid.uuid4()随机生成唯一id，组成excel地址

    def move_mouse(self, page):
        for i in range(3):
            x = random.randint(400, 1000)
            y = random.randint(400, 1000)
            page.mouse.move(x, y)  # 鼠标移动到网页的随机地点
            time.sleep(random.random())  # 随机停留一段时间

    def remove_alert(self):
        """
        去除广告
        :return:
        """
        self.page.wait_for_load_state(state="networkidle")
        if self.page.is_visible('//button/span[text()="我知道了"]'):
            self.page.click('//button/span[text()="我知道了"]')
        if self.page.is_visible('//div[@class="close-box"]'):
            self.page.click('//div[@class="close-box"]')
        time.sleep(1)

    def login(self, playwright):
        for i in range(5):
            try:
                self.browser = playwright.chromium.launch(headless=False,
                                                          # executable_path=ex_path
                                                          )
                self.context = self.browser.new_context()
                self.page = self.context.new_page()
                self.page.goto(setting.get("url"))
                self.page.fill('//input[@id="e2e-login-username"]', setting.get('account'))
                self.page.fill('//input[@id="e2e-login-password"]', setting.get('pwd'))
                with self.page.expect_navigation():  # 捕捉到信息，点击登录按钮
                    self.page.click('//button[@id="e2e-login-submit"]')
                self.page.wait_for_load_state(state="networkidle")  # 等待
                self.page.goto('https://www.chanmama.com/liveRank/todaySales?big_category=&first_category=&second_category=&star_category=')
                self.remove_alert()  # 手写函数，用来移除广告，提示窗
                self.page.wait_for_selector('//div[text()="直播销售额排序"]', timeout=5000)  # 等待元素出现
                return True
            except Exception as e:  # 捕捉异常，捕捉后重新登录
                print(e)
                logger.info(f'登录失败，正在重试第{i + 1}次')  # 登录失败的信息写入日志
                self.browser.close()
        return False

    def get_data_time(self, room_id):
        """
        获取UV,平均停留
        :param room_id:
        :return:
        """
        url = self.base_url + room_id
        new_page = self.context.new_page()  # 打开新页面
        with new_page.expect_response(
                'https://api-service.chanmama.com/v1/douyin/live/star/room/info') as response_info:
            new_page.goto(url)  # 重定向大屏页面
        response = response_info.value
        resp_json = response.json()  # 转化为json格式
        room_info = resp_json.get("data").get("room")  # 获取信息 get()返回None，不会报错
        conversion_rate_percent = round(room_info.get("user_value"), 2)  # uv  用round保留2位小数
        average_residence = room_info.get("average_residence_time")  # 平均停留
        # conversion_rate = float()
        min = divmod(average_residence, 60)[0]  # 【分钟，秒钟】计算有多少分钟
        second = divmod(average_residence, 60)[1]  # 计算有多少秒
        average = str(min) + '‘' + str(second) + '“'
        self.move_mouse(new_page)  # 鼠标在薪打开页面上随机放
        new_page.close()
        return conversion_rate_percent, average  # 返回用户信息和停留时间

    def get_data(self, data_dict):
        data = {}
        rank = data_dict.get('rank')  # 排行
        room_title = data_dict.get('room_title')  # 直播
        nick_name = data_dict.get('nickname')  # 达人
        volume = data_dict.get("volume")  # 直播销量
        amount = data_dict.get('amount')  # 直播销售额
        score = data_dict.get('score')  # 带货热度
        user_peak = data_dict.get('user_peak')  # 人气峰值
        product_size = data_dict.get('product_size')  # 直播商品数
        follower_count = data_dict.get("follower_count")  # 粉丝数
        room_id = data_dict.get("room_id")  # 房间号
        conversion_rate_percent, average_residence_time = self.get_data_time(room_id)  # 获取UV,平均停留
        data['排名'] = rank
        data['直播'] = room_title
        data['达人'] = nick_name
        data['直播销量'] = volume
        data['直播销售额'] = amount
        data['带货热度'] = score
        data['人气峰值'] = user_peak
        data['直播商品数'] = product_size
        data['粉丝数'] = follower_count
        data['UV值'] = conversion_rate_percent  # 用户信息
        data['平均停留'] = average_residence_time
        self.data.append(data)  # 把函数的data，添加进入类的self.data
        logger.info(f"{nick_name}查询完成")  # 信息写入日志

    def get_response(self, response):
        resp_json = response.json()  # 转化为json的格式
        data_list = resp_json.get('data').get('list')  # 从response获取信息
        for i in data_list:
            self.get_data(i)  # 遍历信息并进行 读取，保存到类的data里
            time.sleep(2)

    def process(self):
        """
        逻辑处理流程
        :return:
        """
        self.page.click('//div[text()="直播销量排序"]')  # 进入直播销量排序页面
        with self.page.expect_response(  # 捕捉页面的response
                'https://api-service.chanmama.com/v1/douyin/live/rank/official/daily?star_category=&big_category=&first_category=&second_category=&order=desc&orderby=amount**'
        ) as response_info:
            self.page.click('//div[text()="直播销售额排序"]')
        response = response_info.value  # 捕捉到的response赋值给变量
        self.get_response(response)  # 对response进行处理
        # 翻页 先获取第一页的response，再获取下一页的response
        while self.page.is_enabled('//i[@class="el-icon el-icon-arrow-right"]'):  # 只要下一页存在，持续获取response信息
            with self.page.expect_response(  # 捕捉response信息
                    'https://api-service.chanmama.com/v1/douyin/live/rank/official/daily?star_category=&big_category=&first_category=&second_category=&order=desc&orderby=amount**') as response_info:
                self.page.click('//i[@class="el-icon el-icon-arrow-right"]')  # 点击下一页
            response = response_info.value
            self.get_response(response)

    def write_excel(self):
        """
        生成表格
        :return:
        """
        df = pd.DataFrame(data=self.data)
        df.to_excel(self.result_file, index=False)

    def send_email(self):
        """"
        发送邮件
        """
        recipientAddrs = "zoupengfei@uniner.com;huiling.zhao@leyantech.com"  # 邮件接收地址
        smtpHost = 'smtp.163.com'  # 163服务器地址
        port = 465  # 端口
        sendAddr = '18016454917@163.com'  # 发送人邮箱地址
        password = 'EPNZMCUJPPMCCVRJ'  # 密码 授权码
        msg = email.mime.multipart.MIMEMultipart()  # 加载邮件，进行编写
        msg['from'] = sendAddr  # 发送人邮箱地址
        msg['to'] = recipientAddrs  # 多个收件人的邮箱应该放在字符串中,用字符分隔, 然后用split()分开,不能放在列表中, 因为要使用encode属性
        msg['subject'] = "蝉妈妈数据报表"
        content = "蝉妈妈数据报表"  # 内容
        txt = email.mime.text.MIMEText(content, 'plain', 'utf-8')  # 编写邮件正文(内容)格式
        msg.attach(txt)  # 文件正文(内容)放进去
        logger.info('准备添加附件....')
        part = MIMEApplication(open(self.result_file, 'rb').read())  # 读取之前保存的Excel文件
        file_name = os.path.split(self.result_file)[-1]  # 地址分割，读取列表的最后一个--文件名
        part.add_header('Content-Disposition', 'attachment', filename=file_name)  # 给附件重命名,一般和原文件名一样,改错了可能无法打开.
        msg.attach(part)  # 附件放进去
        logger.info("附件添加成功")  # 输出日志
        smtp = smtplib.SMTP_SSL(smtpHost, port)  # 需要一个安全的连接，用SSL的方式去登录得用SMTP_SSL，之前用的是SMTP（）.端口号465或587
        smtp.login(sendAddr, password)  # 登录发送方的邮箱，和授权码（不是邮箱登录密码）
        smtp.sendmail(sendAddr, recipientAddrs.split(";"), str(msg))  # 注意, 这里的收件方可以是多个邮箱,用";"分开, 也可以用其他符号
        smtp.quit()  # 断开SMTP的链接
        logger.info('邮件发送成功')

    def main(self, playwright):
        login_flg = self.login(playwright)
        if login_flg is False:
            logger.info('登录失败')
        try:
            self.process()
        except Exception:
            pass
        self.write_excel()
        self.send_email()


if __name__ == '__main__':
    with sync_playwright() as playwright:
        chan_ma_ma = ChanMaMa()
        # chan_ma_ma.send_email()
        chan_ma_ma.main(playwright)