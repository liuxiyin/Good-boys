#coding=utf8


"""
乐言-京东商品信息获取
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
import uuid
import pandas as pd
import time
import arrow
import re
import json


ex_path_fire = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
ex_path_fire = ex_path_fire if os.path.exists(ex_path_fire) else None


class JingDong:
    def __init__(self, file):
        self.df = pd.read_excel(file)
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) +'.xlsx'
        self.base_url = "https:"
        self.data = []

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

    def login(self, playwright):
        """
        登录
        :return:
        """
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
        self.page.goto(shop_url)
        self.page.wait_for_load_state(state="networkidle")
        self.page.click('//input[@value="搜本店"]')
        for t in range(60):
            li_objs = self.page.query_selector_all('//div[@class="j-module"]/ul/li')
            if len(li_objs) < 10:
                time.sleep(1)
                continue
            break
        li_list = li_objs[:int(sales_volume)]
        for li in li_list:
            href = li.query_selector('//div[@class="jPic"]/a').get_attribute('href')
            url = self.base_url +href
            url_list.append(url)
        return url_list

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

    def get_number(self, product_details_str):
        """
        获取货号
        :param product_details:
        :return:
        """
        product_dict = {}
        msg_list = product_details_str.split('\n')
        msg_list = [i.replace(' ', '') for i in msg_list if i != ""]
        for item in msg_list:
            try:
                item_list = re.split(r"[:：]", item)
                product_dict[item_list[0]] = item_list[1]
            except Exception:
                continue
        return product_dict

    def get_info(self, page, selector):
        """
        获取数据
        :param page:
        :param seletor:
        :return:
        """
        object = page.query_selector(selector)
        if object:
            return object.text_content()
        return ""

    def get_data_dict(self, page, data_dict):
        """
        获取每一个商品的数据
        :param page:
        :return:
        """
        try:
            ship_title = self.get_info(page, selector='//div[@class="sku-name"]')  # 自定义函数，获取元素文本
            title = self.replace_str(ship_title)  # 文本处理，去除字符
            data_dict['商品标题'] = title
            activity = self.get_info(page, selector='//div[@id="p-ad"]')
            data_dict['活动'] = activity
            price = self.get_info(page, selector='//div[text()="京 东 价"]/following-sibling::div[1]/span[@class="p-price"]/span[contains(@class, "price")]')
            data_dict['秒杀价/京东价'] = price
            coupon_str = self.get_info(page, selector='//div[text()="优 惠 券"]/following-sibling::div[1]/dl/dd')
            coupon = self.replace_str(coupon_str)
            data_dict['优惠券'] = coupon
            promotion = self.get_info(page,selector='//div[@class="prom-item"]/em[@class="hl_red_bg"]')
            data_dict['促销类型'] = promotion
            promotion_detail = self.get_info(page, selector='//div[@class="prom-item"]/em[@class="hl_red"]')
            data_dict['促销详情'] = promotion_detail
            evaluate_no = page.query_selector('//li[contains(@clstag, "shangpin|keycount|product|shangpinpingjia")]/s').text_content()
            data_dict['累计评价'] = evaluate_no
            product_details_str = self.get_info(page, selector='//ul[@class="parameter2 p-parameter-list"]')
            product_details = self.replace_str(product_details_str)
            data_dict['商品详情'] = product_details
            brand = page.wait_for_selector('//ul[@id="parameter-brand"]/li').get_attribute('title')
            data_dict['品牌'] = brand
            number = self.get_number(product_details_str).get('货号')  # 字符串处理，截取货号
            data_dict['货号'] = number
            data_dict['获取情况'] = '成功'
        except Exception as e:
            logger.debug(repr(e))
            data_dict['获取情况'] = '失败'
        finally:
            data_dict['获取时间'] = arrow.now().format("YYYY-MM-DD HH:mm:ss")
            return data_dict

    def get_shop_item_info(self, url_list, shop_url):
        """
        获取店铺商品的信息
        :param url_list:
        :return:
        """
        for url in url_list:
            data_dict = {}
            page = self.context.new_page()
            page.goto(url)
            data_dict['商品链接'] = url
            ship_id = url.split('/')[-1].replace('.html', "")
            data_dict['商品ID'] = ship_id
            data_dict['店铺链接'] = shop_url
            data_dict = self.get_data_dict(page, data_dict)
            self.data.append(data_dict)
            page.close()

    def process(self):
        for index, row in self.df.iterrows():
            shop_url = row['店铺链接']
            sales_volume = row['销量排名前X']
            url_list = self.get_page_url(shop_url, sales_volume)
            self.get_shop_item_info(url_list, shop_url)
            logger.info(f'第{ index + 1}个店铺获取成功')

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
        msg['subject'] = "京东店铺商品获取"
        content = "京东店铺商品获取"  # 内容
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
        if os.path.exists(os.getcwd() + '\\cookies\\') is False:
            os.mkdir(os.getcwd() + '\\cookies\\')
        log_flg = self.login(playwright)
        if not log_flg:
            logger.info('登录京东失败，请重新启动机器人')
            return
        self.process()
        self.browser.close()
        df = pd.DataFrame(data=self.data)
        df.to_excel(self.result_file, index=False)
        if setting.get("recipientAddrs"):  # 从字典setting中取得邮件地址
            self.send_email()
        self.updata_file(strs=self.result_file)


if __name__ == '__main__':
    with sync_playwright() as playwright:
        jing_dong = JingDong(data_file)
        jing_dong.main(playwright)
        # jing_dong.updata_file(jing_dong.result_file)