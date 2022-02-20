"""
拼多多订单解密
"""


from playwright.sync_api import sync_playwright
from email.mime.application import MIMEApplication
from loguru import logger
from Crypto.Hash import MD5
import pandas as pd
import email
import smtplib
import email.mime.multipart
import email.mime.text
import requests
import time
import json
import re
import os
import uuid
import base64


executable_path1 = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
executable_path1 = executable_path1 if os.path.exists(executable_path1) else None

robot_secret = "in00098"
serial_no = "in00098"
base_url = "https://cloud.uniner.com/naas"


setting = {}


class PingDuoDuo:

    def __init__(self, file):
        self.file = file
        self.cookie_dir = os.getcwd() + '\\cookie\\'
        self.data = []
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) + '.xlsx'

    def wait(self, page, second=1):
        """
        等待
        :param page:
        :return:
        """
        try:
            page.wait_for_load_state(state="networkidle", timeout=10000)
            time.sleep(second)
        except:
            return

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

    def remove_advertising(self, page):
        """
        删除广告
        :param page:
        :return:
        """
        if page.is_visible('//span[text()="已安装去使用"]'):
            page.click('//span[text()="已安装去使用"]')
        if page.is_visible('//div[@class="modal-system-notice-modal-title"]'):
            page.click('//span[text()="关闭"]')
        if page.is_visible('//div[contains(@class, "ImportantList_msgbox-header")]'):
            page.click('//i[contains(@class, "ImportantList_close__")]')
        if page.is_visible('//div[@class="modal-system-poster-modal-component-closeIcon"]'):
            page.click('//div[@class="modal-system-poster-modal-component-closeIcon"]')
        if page.is_visible('//div[@class="thirdpart-modal__activity-hide"]'):
            page.click('//div[@class="thirdpart-modal__activity-hide"]')
        if page.is_visible('//span[text()="暂不处理"]'):
            page.click('//span[text()="暂不处理"]')

    def login(self, playwright):
        if os.path.exists(self.cookie_dir) is False:
            os.mkdir(self.cookie_dir)
        logger.info("开始登录拼多多后台")
        browser = playwright.chromium.connect_over_cdp("http://localhost:9002")
        self.context = browser.new_context()
        self.page = self.context.new_page()
        self.load_cookies(path="pinduoduo")  # 自定义函数，如果存在cookies，则读取它
        self.page.goto('https://mms.pinduoduo.com/home/')
        self.remove_advertising(self.page)  # 自定义函数，去除广告
        for t in range(300, 0, -1):
            time.sleep(1)
            self.page.reload()
            self.remove_advertising(self.page)
            self.wait(self.page)
            if self.page.url == "https://mms.pinduoduo.com/home/":
                logger.info(' - 登陆成功')
                self.save_cookies(path="pinduoduo")
                return True
            if t == 300:
                b_img = self.page.wait_for_selector('//div[@class="qr-code"]').screenshot()
                img_base64 = str(base64.b64encode(b_img))
                img_base64 = re.search("b'(.+)'", img_base64).group(1)
                logger.debug(f'data:image/png;base64,{img_base64}')
                img = f'<img class="qrcode-img" src="data:image/png;base64,{img_base64}" style="width: 240px;height: 240px;">'
                logger.info(' - 请打开拼多多商家版APP,使用扫码登录。。。')
                logger.info(img)

            if self.page.is_visible('//p[text()="二维码已失效"]'):
                self.page.click('//button/span[text()="点击刷新"]')
                self.wait(self.page)
                b_img = self.page.wait_for_selector('//div[@class="qr-code"]').screenshot()
                img_base64 = str(base64.b64encode(b_img))
                img_base64 = re.search("b'(.+)'", img_base64).group(1)
                logger.debug(f'data:image/png;base64,{img_base64}')
                img = f'<img class="qrcode-img" src="data:image/png;base64,{img_base64}" style="width: 240px;height: 240px;">'
                logger.info(' - 请打开拼多多商家版APP,使用扫码登录。。。')
                logger.info(img)

            if t % 20 == 0:
                b_img = self.page.wait_for_selector('//div[@class="qr-code"]').screenshot()
                img_base64 = str(base64.b64encode(b_img))
                img_base64 = re.search("b'(.+)'", img_base64).group(1)
                logger.info(f'任务完成, 请下载<a href="https://www.baidu.com" target="_blank">拼多多订单备注</a>，查看格式是否正确')

                # https://www.uniner.com/static/image/home/home_card4.png
                # logger.info('<img class="qrcode-img" src="https://www.uniner.com/static/image/home/home_card4.png" style="width: 240px;height: 240px;">')
                logger.debug(f'data:image/png;base64,{img_base64}')
                img = f"""
                       <img class="qrcode-img" src="data:image/png;base64,{img_base64}" style="width: 240px;height: 240px;">
                       """
                logger.info(' - 请打开拼多多商家版APP,使用扫码登录。。。')
                logger.info(img)

            if t % 5 == 0:
                logger.info(' - 登陆时间还剩 %s 秒' % t)

        else:
            logger.info(' - 登陆超时，请重新再试！')
            return False

    def get_name(self):
        """
        获取收货人姓名
        :return:
        """
        with self.page.expect_response("https://mms.pinduoduo.com/fopen/order/receiver") as response_info:
            self.page.click('//div/a/span[text()="查看"]')
        response = response_info.value
        resp_json = json.loads(response.text())
        name = resp_json.get("result").get("name")
        return name

    def order_decrypt(self, row):
        """
        添加备注
        :param order_id:订单编号
        :return:
        """
        data_dict = {}
        logger.info(f'当前正在处理{row["订单编号"]}')
        data_dict['订单编号'] = row["订单编号"]
        self.page.fill('//input[@placeholder="请输入完整订单编号"]', str(row["订单编号"]))
        with self.page.expect_response('https://mms.pinduoduo.com/mangkhut/mms/recentOrderList', timeout=5000) as response_info:
            self.page.click('//span[text()="查询"]')
        response = response_info.value
        resp_json = json.loads(response.text())
        order_num = resp_json.get("result").get("pageItems")
        if len(order_num) == 0:
            msg = "查无此单号"
            data_dict['备注'] = msg
            self.data.append(data_dict)
            return
        if len(order_num) > 1:
            msg = "该单号有多条订单"
            data_dict['备注'] = msg
            self.data.append(data_dict)
            return
        name = self.get_name()
        data_dict['姓名'] = name
        self.page.wait_for_selector('//a/span[text()="查看手机号"]').click()
        time.sleep(1)
        mobile = self.page.wait_for_selector(
            '//tbody[@data-testid="beast-core-table-middle-tbody"]/tr[@data-testid="beast-core-table-body-tr"]/td[6]/div/div[2]/span').text_content()
        data_dict['手机号码'] = mobile
        self.data.append(data_dict)

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

    def process(self):
        try:
            self.page.goto('https://mms.pinduoduo.com/orders/list', timeout=10000)
        except Exception:
            pass
        self.wait(self.page)
        self.remove_advertising(self.page)
        for index, row in self.df.iterrows():
            if pd.isna(row['订单编号']):
                continue
            logger.info(f'当前执行{index + 1}条订单, 剩余{len(self.df) - (index + 1 )}条')
            try:
                self.order_decrypt(row)
            except Exception:
                logger.info(f'第{index + 1}条处理失败')
                continue
            finally:
                time.sleep(5)
        df = pd.DataFrame(data=self.data)
        df.to_excel(self.result_file, index=False)

    def main(self, playwright):
        try:
            self.df = pd.read_excel(self.file)
            log_flg = self.login(playwright)
            if log_flg is False:
                return
            self.process()
        except Exception as e:
            logger.info(e)
        finally:
            self.page.close()
            if setting.get("recipientAddrs"):
                self.send_email()
            self.updata_file(strs=self.result_file)


if __name__ == '__main__':
    import subprocess
    server = subprocess.Popen("./ms-playwright/chromium-907428/chrome-win/chrome --remote-debugging-port=9002")
    time.sleep(2)
    data_file = r"C:\Users\EDZ\Desktop\新建文件夹\拼多多订单解密.xlsx"
    with sync_playwright() as playwright:
        ping_dd = PingDuoDuo(data_file)
        ping_dd.main(playwright)
    server.kill()