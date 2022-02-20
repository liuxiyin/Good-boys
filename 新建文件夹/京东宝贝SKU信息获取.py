"""
京东宝贝SKU信息获取
"""


from playwright.sync_api import sync_playwright
from loguru import logger
from email.mime.application import MIMEApplication
from Crypto.Hash import MD5
import email.mime.multipart
import email.mime.text
import requests
import smtplib
import os
import json
import uuid
import pandas as pd
import time
import arrow
import re


ex_path_fire = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
ex_path_fire = ex_path_fire if os.path.exists(ex_path_fire) else None


class JingDong:

    def __init__(self, file):
        self.df = pd.read_excel(file)
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) + '.xlsx'
        self.base_url = "https:"
        self.data = []
        self.cook_dir = os.getcwd() + '\\cookies\\'

    def save_cookies(self,path="cookies"):
        self.context.storage_state(path=f'cookies/{path}.json')

    def load_cookies(self,path="cookies"):
        if os.path.exists(f'cookies/{path}.json'):
            with open(f'cookies/{path}.json') as f:
                json_str = json.load(f)
                self.context.add_cookies(json_str['cookies'])
            return True
        else:
            return False

    def print_img(self):
        """
        打印二维码
        :return:
        """
        try:
            img_64 = self.page.wait_for_selector('//div[@class="qrcode-img"]', timeout=3000).screenshot()
            img = f""" <img class="qr" src="{img_64}" alt="" style="width: 180px;height: 180px;"> """
            logger.info(img)
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
                        self.save_cookies(path='jingdong')  # 保存当前页面的cookies
                        return True
                    if t % 30 == 0:
                        logger.info('暂未登录，请扫码登录')
                        self.page.click('//a[text()=" 扫码登录"]')
                        self.print_img()
                    if self.page.is_visible('//p[text()="二维码已失效"]'):
                        self.page.click('//a[@class="refresh-btn"]')
                raise Exception("登录超时")
            except Exception as e:
                logger.info(repr(e))
                time.sleep(1)
        return False

    def process_color_price(self, data_dict1, data_dict2):
        """
        获取每一个颜色的价格
        :param color_obj_list:
        :param data:
        :return:
        """
        color_obj_list = self.page.query_selector_all('//div[@data-type="颜色"]/div[@class="dd"]/div')
        sku_list = [i.get_attribute("data-sku") for i in color_obj_list]
        for i in sku_list:
            try:
                item = self.page.wait_for_selector(f'//div[@data-type="颜色"]/div[@class="dd"]/div[@data-sku="{i}"]')
                item.click()
                time.sleep(1)
                color_obj = self.page.query_selector(f'//div[@data-type="颜色"]/div[@class="dd"]/div[@data-sku="{i}"]')
                if color_obj:
                    data_dict2['颜色'] = color_obj.get_attribute('title')
                else:
                    data_dict2['颜色'] = ""
                price = self.page.query_selector('//span[@class="p-price"]/span[contains(@class, "price")]').text_content()
                data_dict2['价格'] = price
                data = {**data_dict1, **data_dict2}
                self.data.append(data)
            except Exception as e:
                logger.info(repr(e))
                continue

    def get_shop_sku_data(self):
        """
        sku信息
        :return:
        """
        data_dict = {}
        data_dict['商品链接'] = self.page.url
        ship_id = self.page.url.split('/')[-1].replace('.html', "")
        data_dict['商品ID'] = ship_id
        src = self.base_url + self.page.wait_for_selector('//img[@id="spec-img"]').get_attribute('src')
        data_dict['SKU图'] = src
        size_obj_list = self.page.query_selector_all('//div[@data-type="尺码"]/div[@class="dd"]/div')
        size_val_list = [i.get_attribute("data-value") for i in size_obj_list]
        for item in size_val_list:
            size_obj = self.page.query_selector(f'//div[@data-type="尺码"]/div[@class="dd"]/div[@data-value={item}]')
            data_dict_sun = {}
            data_dict_sun['尺码'] = item
            size_obj.click()
            self.page.wait_for_load_state(state="networkidle")
            self.process_color_price(data_dict, data_dict_sun)
            time.sleep(1)

    def process(self):
        """
        流程处理
        :return:
        """
        for index, row in self.df.iterrows():
            shop_url = row['商品链接']
            self.page.goto(shop_url)
            self.page.wait_for_load_state(state="networkidle")
            self.get_shop_sku_data()  # 获取商品sku信息
            logger.info(f'第{index +1}个商品获取完毕')

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
        msg['subject'] = "京东sku数据报表"
        content = "京东sku数据报表"  # 内容
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

    def main(self, playwright):
        login_flg = self.login(playwright)
        if not login_flg:
            logger.info('登录失败, 请重新启动机器人')
            return
        self.process()
        self.browser.close()
        logger.info(self.data)
        df = pd.DataFrame(data=self.data)
        df.to_excel(self.result_file, index=False)
        if setting.get("recipientAddrs"):
            self.send_email()
        self.updata_file(strs=self.result_file)


if __name__ == '__main__':
    data_file = r"C:\Users\EDZ\Desktop\乐言-京东宝贝SKU信息获取(1).xlsx"
    with sync_playwright() as playwright:
        jing_dong = JingDong(data_file)
        jing_dong.main(playwright)