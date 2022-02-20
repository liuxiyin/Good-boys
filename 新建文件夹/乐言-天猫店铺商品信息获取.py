"""
乐言-天猫店铺商品信息获取
"""
import base64

from playwright.sync_api import sync_playwright
from loguru import logger
from email.mime.application import MIMEApplication
from Crypto.Hash import MD5
import requests
import email.mime.multipart
import email.mime.text
import smtplib
import os
import uuid
import pandas as pd
import time
import arrow
import re
import json


ex_path = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
ex_path = ex_path if os.path.exists(ex_path) else None


class TianMao:
    def __init__(self, file):
        self.df = pd.read_excel(file)
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) +'.xlsx'
        self.base_url = "https:"
        self.data = []
        self.columns = ["店铺网址","商品链接", "商品ID", "商品标题", "商品小标题", "原始价格", "活动价格", "活动标签", "月销量", "累计评价", "品牌名称",
                        "产品参数", "货/款号", "库存总数", "获取情况", "获取时间"]

    def save_cookies(self,path="cookies"):
        if os.path.exists('cookies') is False:
            os.mkdir('cookies')
        self.context.storage_state(path=f'cookies/{path}.json')

    def load_cookies(self,path="cookies"):
        if os.path.exists(f'cookies/{path}.json'):
            with open(f'cookies/{path}.json') as f:
                json_str = json.load(f)
                self.context.add_cookies(json_str['cookies'])
            return True
        else:
            return False

    def print_img(self,page):
        """
        前端显示二维码
        :return:
        """
        page.wait_for_selector('//div[@class="qrcode-img"]').screenshot(path='data/1.png')
        with open('data/1.png', 'rb') as f:
            img_64 = str(base64.b64encode(f.read()))
        img_64 = re.search("b'(.+)'", img_64).group(1)
        img = f""" <img class="qr" src="data:image/png;base64,{img_64}" alt="" style="width: 180px;height: 180px;"> """
        logger.info(img)
        logger.info('请打开淘宝APP进行扫码登录')

    def click_err_img(self, page):
        """
        点击二维码刷新
        :param page:
        :return:
        """
        try:
            if page.is_visible('//div[@class="qrcode-error"]', timeout=3000):
                page.click('//button[@class="refresh"]')
        except Exception:
            return

    def login(self, playwright):
        try:
            # self.browser = playwright.chromium.connect_over_cdp("http://localhost:9002")
            self.browser = playwright.chromium.launch(headless=False, executable_path=ex_path)
            self.context = self.browser.new_context()
            self.page = self.context.new_page()
            # self.load_cookies(path="tianmao")
            self.page.goto('https://www.tmall.com/')
            # self.page.goto('https://i.taobao.com/my_taobao.htm?spm')
            for t in range(300, 0, -1):
                time.sleep(1)
                try:
                    user_name = self.page.query_selector('//p[@id="login-info"]').text_content()
                    if "请登录" not in user_name:
                        logger.info('天猫商城已登录')
                        self.save_cookies(path="tianmao")
                        return True
                except Exception:
                    pass
                if t % 30 == 0:
                    self.page.goto('https://login.tmall.com/')
                    self.page.wait_for_load_state('networkidle')
                    self.login_frame = self.page.wait_for_selector('//iframe[@id="J_loginIframe"]').content_frame()
                    self.login_frame.click('//i[@class="iconfont icon-qrcode"]')
                    self.print_img(self.login_frame)  # 二维码输出到日志
                self.click_err_img(self.login_frame)
            raise Exception("登录失败")  # 循环完毕后，未登录成功，再执行
        except Exception as e:
            logger.info(repr(e))
            self.browser.close()
            return False

    def get_page_url(self, shop_url, sales_volume):
        """
        获取url
        :param shop_name:
        :return:
        """
        url_list = []
        url = shop_url +'search.htm'
        self.page.goto(url)
        logger.info('正在进入店铺')
        self.page.wait_for_load_state(state="networkidle")
        dl_objs = self.page.query_selector_all('//dl')
        dl_list = dl_objs[:int(sales_volume)]
        for dl in dl_list:
            href = dl.query_selector('//dd[@class="detail"]/a').get_attribute('href')
            url = self.base_url +href
            url_list.append(url)
        return url_list

    def get_activity_label(self, page):
        """
        获取活动标签
        :param page:
        :return:
        """
        activity_label = page.query_selector('//div[@class="tb-skin"]/p')
        if activity_label:
            if "聚划算" in activity_label.text_content():
                return "聚划算"
        return ""

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

    def get_msg_item(self, page):
        """
        获取每一个商品的信息
        :param page:
        :return:
        """
        data_dict = {i: "" for i in self.columns}  # 字典读取列名
        try:
            data_dict["商品链接"] = page.url
            ship_id = re.findall(r".*id=(.+?)&.*", page.url)[0]
            data_dict['商品ID'] = ship_id
            ship_title_str = page.wait_for_selector('//div[@class="tb-detail-hd"]/h1').text_content()
            ship_title = self.replace_str(ship_title_str)
            data_dict['商品标题'] = ship_title
            small_title_str = page.wait_for_selector('//div[@class="tb-detail-hd"]/p').text_content()
            small_title = self.replace_str(small_title_str)
            data_dict['商品小标题'] = small_title
            old_price = page.wait_for_selector('//dt[text()="价格"]/following-sibling::dd[1]/span[@class="tm-price"]').text_content()
            data_dict['原始价格'] = old_price
            activity_price = page.wait_for_selector('//dt[text()="活动价"]/following-sibling::dd[1]/span[@class="tm-price"]').text_content()
            data_dict['活动价格'] = activity_price
            activity_label = self.get_activity_label(page)
            data_dict['活动标签'] = activity_label
            monthly_sales = page.wait_for_selector('//li[@data-label="月销量"]/div/span[@class="tm-count"]').text_content()
            data_dict["月销量"] = monthly_sales
            cumulative_evaluation = page.wait_for_selector('//span[text()="累计评价"]/following-sibling::span[@class="tm-count"]').text_content()
            data_dict['累计评价'] = cumulative_evaluation
            product_parameters_str = page.wait_for_selector('//ul[@id="J_AttrUL"]').text_content()
            product_parameters = self.replace_str(product_parameters_str)
            data_dict['产品参数'] = product_parameters
            ship_no = re.findall(r".*货号:(.+?)是否商场.*", product_parameters)[0]
            shop_name = re.findall(r".*品牌:(.+?)尺码", product_parameters)[0]
            data_dict['品牌名称'] = shop_name
            data_dict["货/款号"] = ship_no
            stock_str = page.wait_for_selector('//em[@id="J_EmStock"]').text_content()
            stock_number = re.findall(r"\d+", stock_str)[0]
            data_dict['库存总数'] = stock_number
            data_dict["获取情况"] = "成功"
        except Exception as e:
            logger.info(repr(e))
            data_dict["获取情况"] = "失败"
        finally:
            data_dict['获取时间'] = arrow.now().format('YYYY-MM-DD')
            return data_dict

    def get_shop_item_info(self, url_list):
        """
        获取店铺商品的信息
        :param url_list:
        :return:
        """
        for url in url_list:
            page = self.context.new_page()
            page.goto(url)
            data_dict = self.get_msg_item(page)  # 获取每一个商品的信息
            return data_dict

    def process(self):
        for index, row in self.df.iterrows():
            shop_url = row['店铺链接']
            sales_volume = row['销量排名前X']
            url_list = self.get_page_url(shop_url, sales_volume)
            data_dict = self.get_shop_item_info(url_list)
            data_dict["店铺链接"] = shop_url
            self.data.append(data_dict)

    def updata_file(self, strs, nass_name='文件'):
        obj = MD5.new()
        obj.update(robot_secret.encode("utf-8"))  # gb2312 Or utf-8
        robotPwd = obj.hexdigest()

        url = f'{base_url}/file/upload/{serial_no}/{robotPwd}'
        header = {
            # "content-type": "application/json"
        }
        try:
            data = {}
            files = {'file': open(strs, 'rb')}
            r = requests.post(url=url, headers=header, files=files)
            ret = r.json().get('data')
            url = ret.get('url').replace('文件下载', nass_name)
            path = ret.get('path')
            logger.success(url, filename=path)
        except Exception as e:
            print(e)

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
        msg['subject'] = "数据报表"
        content = "天猫店铺数据报表"  # 内容
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

    def main(self, playwright):
        self.login(playwright)
        self.process()
        self.browser.close()
        df = pd.DataFrame(data=self.data)
        df.to_excel(self.result_file, index=False)
        self.updata_file(strs=self.result_file)
        self.send_email()
        logger.info('任务完成，请在邮箱中下载任务文件')


if __name__ == '__main__':
    # import subprocess
    # server = subprocess.Popen("./ms-playwright/chromium-901522/chrome-win/chrome --remote-debugging-port=9002")
    # time.sleep(2)
    # data_file = r"C:\Users\EDZ\Desktop\输入.xlsx"
    with sync_playwright() as playwright:
        try:
            tian_mao = TianMao(data_file)
            tian_mao.main(playwright)
        except Exception as e:
            logger.info(e)
        # finally:
        #     server.kill()
