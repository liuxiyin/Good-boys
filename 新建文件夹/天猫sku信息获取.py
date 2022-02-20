import base64
from playwright.sync_api import sync_playwright
from email.mime.application import MIMEApplication
from loguru import logger
from Crypto.Hash import MD5
import pandas as pd
import os
import uuid
import requests
import re
import json
import time
import email.mime.multipart
import email.mime.text
import smtplib


class TianMaoSKU:
    def __init__(self, file):
        self.df = pd.read_excel(file)
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) + '.xlsx'
        self.data = []
        self.columns =["商品链接", "商品ID", "SKU图", "尺码", "颜色分类", "库存量", "价格"]
        self.base_url = "https://"

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
            self.browser = playwright.chromium.connect_over_cdp("http://localhost:9002")

            self.context = self.browser.new_context()
            self.page = self.context.new_page()
            self.load_cookies(path="tianmao")
            self.page.goto('https://www.tmall.com/')
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
                    self.login_frame = self.page.wait_for_selector('//iframe[@id="J_loginIframe"]').content_frame()
                    self.login_frame.click('//i[@class="iconfont icon-qrcode"]')
                    self.print_img(self.login_frame)
                self.click_err_img(self.login_frame)
            raise Exception("登录失败")
        except Exception as e:
            logger.info(repr(e))
            self.browser.close()
            return False

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

    def get_size_list(self, li_objs):
        """
        获取尺码列表
        :param li_objs:
        :return:
        """
        size_list = []
        for i in li_objs:
            text = i.text_content()
            size = re.findall(r"\d+", self.replace_str(text))[0]
            size_list.append(size)
        return size_list

    def process_size(self, data_dict1, data_dict2):
        """
        处理每一条size
        :return:
        """
        color_list = self.page.query_selector_all('//ul[@data-property="颜色分类"]/li')
        color_list_obj = [i for i in color_list if i.is_visible()]
        for i in color_list_obj:
            color = i.get_attribute('title')
            data_dict2['颜色分类'] = color
            if i.get_attribute('class') is None or 'selected' not in i.get_attribute('class'):
                i.click()
                time.sleep(0.1)
            old_price = self.page.wait_for_selector('//dt[text()="价格"]/following-sibling::dd[1]/span[@class="tm-price"]').text_content()
            data_dict2['价格'] = old_price
            if i.get_attribute('class') == "tb-out-of-stock":
                stock_number = "0"
            else:
                stock_str = self.page.wait_for_selector('//em[@id="J_EmStock"]').text_content()
                stock_number = re.findall(r"\d+", stock_str)[0]
            data_dict2['库存量'] = stock_number
            res = {**data_dict1, **data_dict2}
            self.data.append(res)
            i.click()
            time.sleep(0.3)

    def get_data(self, url):
        """
        获取信息
        :return:
        """
        data_dict = {i: "" for i in self.columns}
        data_dict['商品链接'] = url
        ship_id = re.findall(r".*id=(.+?)&.*", self.page.url)[0]
        data_dict['商品ID'] = ship_id
        src = self.page.wait_for_selector('//img[@id="J_ImgBooth"]').get_attribute('src')
        img_url = self.base_url + src
        data_dict['SKU图'] = img_url
        li_objs = self.page.query_selector_all('//div[@class="tb-sku"]/dl[@class="tb-prop tm-sale-prop tm-clear "]/dd/ul/li')
        li_objs = [i for i in li_objs if i.is_visible()]
        size_list = self.get_size_list(li_objs)
        for index, item in enumerate(li_objs):
            data_size = {}
            if item.get_attribute('class') is None or 'selected' not in item.get_attribute('class'):
                item.click()
                time.sleep(0.1)
            data_size['尺码'] = size_list[index]
            self.process_size(data_dict, data_size)
            time.sleep(1)

    def process(self):
        for index, row in self.df.iterrows():
            url = row['商品链接']
            logger.info(f'正在获取第{index + 1}个商品')
            try:
                self.page.goto(url, timeout=10000)
            except Exception:
                pass
            self.get_data(url)
            logger.info(f'第{index + 1}个商品获取完毕')

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
        msg['subject'] = "天猫sku数据报表"
        content = "天猫sku数据报表"  # 内容
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
        log_flg = self.login(playwright)
        if log_flg is False:
            logger.info('登录失败，请重启机器人')
            return
        self.process()
        self.browser.close()
        df = pd.DataFrame(data=self.data)
        df.to_excel(self.result_file, index=False)
        self.updata_file(strs=self.result_file)
        self.send_email()
        logger.info('任务完成,请查收邮箱文件')


if __name__ == '__main__':
    import subprocess
    server = subprocess.Popen("ms-playwright/chromium-901522/chrome-win/chrome --remote-debugging-port=9002 --")
    time.sleep(2)
    # data_file = r"C:\Users\EDZ\Desktop\乐言-京东宝贝SKU信息获取(1).xlsx"
    with sync_playwright() as playwright:
        try:
            sku = TianMaoSKU(data_file)
            sku.main(playwright)
        except Exception as e:
            logger.info(e)
        finally:
            server.kill()

