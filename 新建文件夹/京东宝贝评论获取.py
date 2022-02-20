"""
京东宝贝评论获取
"""


from playwright.sync_api import sync_playwright
from loguru import logger
from email.mime.application import MIMEApplication
from Crypto.Hash import MD5
import email.mime.multipart
import email.mime.text
import smtplib
import requests
import os
import json
import uuid
import pandas as pd
import time


ex_path_fire = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
ex_path_fire = ex_path_fire if os.path.exists(ex_path_fire) else None


class JingDong:
    def __init__(self, file):
        self.df = pd.read_excel(file)
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) +'.xlsx'
        self.base_url = "https:"
        self.shop_url = "https://search.jd.com/Search?keyword="
        self.data = []

    def save_cookies(self, path="cookies"):
        self.context.storage_state(path=f'cookies/{path}.json')  # 将当前浏览器上下文的全部状态（cookies信息）保存下来

    def load_cookies(self, path="cookies"):
        if os.path.exists(f'cookies/{path}.json'):  # 如果存在浏览器状态的信息
            with open(f'cookies/{path}.json') as f:
                json_str = json.load(f)  # 用json读取保存cookies的文件
                self.context.add_cookies(json_str['cookies'])  # 向页面添加cookies信息
            return True
        else:
            return False

    def print_img(self):
        """
        打印二维码
        :return:
        """
        try:
            img_64 = self.page.wait_for_selector('//div[@class="qrcode-img"]', timeout=3000).screenshot()  # 把二维码截图
            img = f""" <img class="qr" src="{img_64}" alt="" style="width: 180px;height: 180px;"> """  # 设定二维码显示大小
            logger.info(img)  # 输出二维码到日志
            logger.info('请打开京东APP进行扫码登录')
        except Exception:
            return

    def replace_str(self, str):
        """
        去除特殊字符
        :param str:
        :return:
        """
        spec_list = ['\n', '\t', '\xa0', ' ']
        for i in spec_list:
            str = str.replace(i, '')
        return str

    def login(self, playwright):
        """
        登录
        :return:
        """
        for i in range(5):
            try:
                self.browser = playwright.chromium.launch(headless=False,
                                                          executable_path=ex_path_fire
                                                          )
                self.context = self.browser.new_context()

                self.page = self.context.new_page()
                # self.page.goto("https://passport.jd.com/new/login.aspx")
                self.load_cookies(path="jingdong")
                self.page.goto('https://www.jd.com/')
                #
                # self.print_img()
                for t in range(300, 0, -1):
                    time.sleep(1)
                    if self.page.url == "https://www.jd.com/":
                        logger.info('京东登录成功')
                        self.save_cookies(path='jingdong')
                        return True
                    if t % 30 == 0:
                        logger.info('暂未登录，请扫码登录')
                        self.page.click('//a[text()=" 扫码登录"]')
                        self.print_img()
                    if self.page.is_visible('//p[text()="二维码已失效"]'):
                        self.page.click('//a[@class="refresh-btn"]')
                raise Exception("登录超时")
            except Exception as e:
                logger.info(repr(e))  # 把异常信息当做日志输出
                time.sleep(1)
        return False

    def get_url_list(self, sales_volume):  # 获取商店前几商品的网络地址
        url_list = []
        self.page.click('//input[@value="搜本店"]')
        self.page.wait_for_load_state(state="networkidle")
        for i in range(60):
            li_list = self.page.query_selector_all('//div[@class="jSearchListArea"]/div[@class="j-module"]/ul/li')  # 获得所有元素作为list
            if len(li_list) > int(sales_volume):
                break
            time.sleep(1)
        for li in li_list[:int(sales_volume)]:
            href = li.query_selector('//div[@class="jPic"]/a').get_attribute('href')
            url = self.base_url + href
            url_list.append(url)
        return url_list

    def get_info(self, content, selector):
        """
        获取具体信息
        :param selector:
        :return:
        """
        object1 = content.query_selector(selector)
        if object1:
            return object1.text_content()  # 获取该元素的信息
        return ""

    def get_self_content(self, page, shop_url):
        """
        获取评论信息
        :param page:
        :return:
        """
        data = {}
        page.click('//li[contains(text(),"商品评价")]')  # 进入商品评价页面
        time.sleep(1)
        ship_title = self.get_info(page, selector='//div[@class="sku-name"]')  # 商品标题
        title = self.replace_str(ship_title)  # 用自定义函数去除特殊符号
        data['商品标题'] = title  # 把商品标题放入本地 字典data 中
        data['商品链接'] = page.url
        ship_id = page.url.split('/')[-1].replace('.html', "")  # page.url切片并替换，获得ship_id商品ID
        data['商品ID'] = ship_id
        data['店铺链接'] = shop_url
        content_list = page.query_selector_all('//div[@id="comment"]//div[@class="comment-item"]')  # 获得本页面该元素全内容
        for content in content_list:
            data_dict = {}
            nick_name_str = self.get_info(content, selector='//div[@class="user-info"]')
            nick_name = self.replace_str(nick_name_str)  # 用自定义函数去除特殊符号
            data_dict['买家昵称'] = nick_name
            con = self.get_info(content, selector='//p[@class="comment-con"]')
            data_dict['评价内容'] = con
            color = self.get_info(content, selector='//div[@class="comment-message"]/div[@class="order-info"]/span[1]')
            data_dict['颜色'] = color
            size = self.get_info(content, selector='//div[@class="comment-message"]/div[@class="order-info"]/span[2]')
            data_dict['尺码'] = size
            date_list = content.query_selector_all('//div[@class="comment-message"]/div[@class="order-info"]/span')
            con_date = self.get_info(content, selector=f'//div[@class="comment-message"]/div[@class="order-info"]/span[{len(date_list)}]')
            data_dict['评价日期'] = con_date  # 读取列表的最后一个作为评价日期
            res = {**data, **data_dict}  # 表示随机长度的字典
            self.data.append(res)

    def get_data(self, url_list, shop_url):  # 获取每个商品的评论
        """
        获取每个商品的评论
        :param url_list:
        :return:
        """
        page = self.context.new_page()
        for url in url_list:
            try:
                page.goto(url, timeout=1000)
            except Exception:
                pass
            time.sleep(1)
            self.get_self_content(page, shop_url)
        page.close()

    def process(self):  # 每个店铺销量前几的商品信息获取
        for index, row in self.df.iterrows():  # iterrows迭代，
            shop_url = row['店铺链接']
            sales_volume = row['销量排名前X']
            self.page.goto(shop_url)
            url_list = self.get_url_list(sales_volume)  # 商品链接
            self.get_data(url_list, shop_url)  # 获取该商店每个商品的评论
            logger.info(f'第{index + 1}个商品获取完毕')
            time.sleep(1)

    def send_email(self):
        """"
        发送邮件
        """
        recipientAddrs = setting.get("recipientAddrs")
        smtpHost = 'smtp.163.com'
        port = 465
        sendAddr = '18016454917@163.com'
        password = 'EPNZMCUJPPMCCVRJ'
        msg = email.mime.multipart.MIMEMultipart()
        msg['from'] = sendAddr  # 发送人邮箱地址
        msg['to'] = recipientAddrs  # 多个收件人的邮箱应该放在字符串中,用字符分隔, 然后用split()分开,不能放在列表中, 因为要使用encode属性
        msg['subject'] = "京东宝贝评论获取"
        content = "京东宝贝评论获取报表"  # 内容
        txt = email.mime.text.MIMEText(content, 'plain', 'utf-8')
        msg.attach(txt)
        logger.info('准备添加附件....')
        part = MIMEApplication(open(self.result_file, 'rb').read())
        file_name = os.path.split(self.result_file)[-1]
        part.add_header('Content-Disposition', 'attachment', filename=file_name)  # 给附件重命名,一般和原文件名一样,改错了可能无法打开.
        msg.attach(part)
        logger.info("附件添加成功")
        smtp = smtplib.SMTP_SSL(smtpHost, port)  # 需要一个安全的连接，用SSL的方式去登录得用SMTP_SSL，之前用的是SMTP（）.端口号465或587
        smtp.login(sendAddr, password)  # 发送方的邮箱，和授权码（不是邮箱登录密码）
        smtp.sendmail(sendAddr, recipientAddrs.split(";"), str(msg))  # 注意, 这里的收件方可以是多个邮箱,用";"分开, 也可以用其他符号
        smtp.quit()
        logger.info('邮件发送成功')

    def updata_file(self, strs1, nass_name='文件'):
        obj = MD5.new()
        obj.update(robot_secret.encode("utf-8"))  # gb2312 Or utf-8
        robotPwd = obj.hexdigest()

        url = f'{self.base_url}/file/upload/{serial_no}/{robot_secret}'
        header = {
            # "content-type": "application/json"
        }
        try:
            data = {}
            files = {'file': open(strs1, 'rb')}
            r = requests.post(url=url, headers=header, files=files)
            ret = r.json().get('data')
            url = ret.get('url').replace('文件下载', nass_name)
            path = ret.get('path')
            logger.success(url, filename=path)
        except Exception as e:
            print(e)

    def main(self):
        log_flg = self.login(playwright)
        if log_flg is False:
            logger.info('登录失败, 请重新启动机器人')
        try:
            self.process()
        except Exception as e:
            logger.info(e)
        self.browser.close()  # 关闭浏览器
        df = pd.DataFrame(data=self.data)  # 获取的data用pandas的dataframe格式打开
        df.to_excel(self.result_file, index=False)
        self.send_email()
        self.updata_file(strs1=self.result_file)


# data_file = r"C:\Users\EDZ\Desktop\新建文件夹\乐言-京东宝贝评论获取.xlsx"
with sync_playwright() as playwright:
    jing_dong = JingDong(data_file)
    jing_dong.main()