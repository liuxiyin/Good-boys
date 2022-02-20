"""
乐言-天猫全部宝贝评论获取
"""
import base64

from playwright.sync_api import sync_playwright
from email.mime.application import MIMEApplication
from loguru import logger
from Crypto.Hash import MD5
import pandas as pd
import os
import requests
import uuid
import json
import re
import time
import email.mime.multipart
import email.mime.text
import smtplib


ex_path_fire = "./ms-playwright/chromium-907428/chrome-win/chrome.exe"
ex_path_fire = ex_path_fire if os.path.exists(ex_path_fire) else None


class TianMao:
    def __init__(self, file):
        self.df = pd.read_excel(file)
        self.df = pd.read_excel(file)
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) + '.xlsx'
        # self.result_file = r"D:\GITFile\发货异常\data\9f4e5855-2d7f-4c54-8d7f-d3e9e4808b6a.xlsx"
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

    def print_img(self,page):
        """
        前端显示二维码
        :return:
        """
        if page:
            page.wait_for_selector('//div[@class="qrcode-img"]').screenshot(path='data/1.png')
            with open('data/1.png', 'rb') as f:
                img_64 = str(base64.b64encode(f.read()))
            img_64 = re.search("b'(.+)'", img_64).group(1)
            img = f""" <img class="qr" src="data:image/png;base64,{img_64}" alt="" style="width: 180px;height: 180px;"> """
            logger.info(img)
            logger.info('请打开淘宝APP进行扫码登录')
        return

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

    def get_page_url(self, shop_url, sales_volume):
        """
        获取url
        :param shop_name:
        :return:
        """
        url_list = []
        url = shop_url +'search.htm'
        self.page.goto(url)
        time.sleep(3)
        dl_objs = self.page.query_selector_all('//div[@id="J_ShopSearchResult"]//dl')
        dl_list = dl_objs[:int(sales_volume)]
        for dl in dl_list:
            href = dl.wait_for_selector('//dd[@class="detail"]/a').get_attribute('href')
            url = self.base_url +href
            url_list.append(url)
        return url_list

    def get_commodity_evaluate(self, page, data_dict):
        """
        获取每一个商品的评价信息
        :param page:
        :return:
        """
        page.wait_for_selector('//a[contains(text(), "累计评价")]').click()
        time.sleep(2)
        tr_list = page.query_selector_all('//div[@class="rate-grid"]/table/tbody/tr')  # 评论列表
        for tr in tr_list:
            evaluate_dict = {}
            content = tr.wait_for_selector('//td[@class="tm-col-master"]//div[@class="tm-rate-fulltxt"]').text_content()
            evaluate_dict['评价内容'] = content
            date = tr.wait_for_selector('//td[@class="tm-col-master"]//div[@class="tm-rate-date"]').text_content()
            evaluate_dict['评价时间'] = date
            color_str = tr.wait_for_selector('//td[@class="col-meta"]/div[@class="rate-sku"]/p[contains(@title, "颜色分类")]').text_content()
            color = color_str.replace('颜色分类：', "")
            evaluate_dict['颜色分类'] = color
            size_str = tr.wait_for_selector('//td[@class="col-meta"]/div[@class="rate-sku"]/p[2]').text_content()
            size = re.findall(r"\d+", size_str)[0]
            evaluate_dict['尺码'] = size
            nick_name = tr.wait_for_selector('//td[@class="col-author"]').text_content()
            evaluate_dict['买家昵称'] = nick_name
            res = {**data_dict, **evaluate_dict}
            self.data.append(res)

    def get_shop_item_info(self, url_list, data):
        """
        获取店铺商品的信息
        :param url_list:
        :return:
        """
        for url in url_list:
            self.page.goto(url)
            ship_id = re.findall(r".*id=(.+?)&.*", url)[0]
            data['商品ID'] = ship_id
            data['商品链接'] = url
            ship_title_str = self.page.wait_for_selector('//div[@class="tb-detail-hd"]/h1').text_content()
            ship_title = self.replace_str(ship_title_str)
            data['商品标题'] = ship_title
            self.get_commodity_evaluate(self.page, data)  # 获取每一个商品的评价信息
            time.sleep(1)

    def process(self):
        """
        业务流程处理
        :return:
        """
        for index, row in self.df.iterrows():
            data_dict = {}
            shop_url = row['店铺链接']
            data_dict['店铺链接'] = shop_url
            sales_volume = row['销量排名前X']
            url_list = self.get_page_url(shop_url, sales_volume)
            self.get_shop_item_info(url_list, data_dict)
            logger.info(f'第{index + 1}条获取评论完毕')

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
        recipientAddrs = setting.get("recipientAddrs")  # 邮件接收地址
        smtpHost = 'smtp.163.com'  # 163邮箱
        port = 465  # 465端口
        sendAddr = '18016454917@163.com'  # 邮件发送地址
        password = 'EPNZMCUJPPMCCVRJ'  # 通行码
        msg = email.mime.multipart.MIMEMultipart()
        msg['from'] = sendAddr  # 发送人邮箱地址
        msg['to'] = recipientAddrs  # 多个收件人的邮箱应该放在字符串中,用字符分隔, 然后用split()分开,不能放在列表中, 因为要使用encode属性
        msg['subject'] = "天猫宝贝评论获取数据报表"
        content = "天猫宝贝评论获取数据报表"  # 内容
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

    def main(self,playwright):
        log_flg = self.login(playwright)
        if not log_flg:
            logger.info('登录失败请重新启动机器人')
            return
        self.process()
        self.browser.close()
        df = pd.DataFrame(data=self.data)
        df.to_excel(self.result_file, index=False)
        self.updata_file(strs=self.result_file)
        self.send_email()
        logger.info('任务完成，请在邮箱中下载任务文件')


if __name__ == '__main__':
    import subprocess
    server = subprocess.Popen("ms-playwright/chromium-907428/chrome-win/chrome --remote-debugging-port=9002 --" )
    time.sleep(2)
    # data_file = "C:\\Users\\EDZ\\Desktop\\乐言-京东宝贝评论获取.xlsx"
    with sync_playwright() as playwright:
        tian_mao = TianMao(data_file)
        tian_mao.main(playwright)
    server.kill()