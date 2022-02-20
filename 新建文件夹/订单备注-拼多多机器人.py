"""
订单备注-拼多多机器人
"""
import email
import smtplib
from email.mime.application import MIMEApplication

from playwright.sync_api import sync_playwright
from loguru import logger
import os
import json
import uuid
import time
import utility
import pandas as pd
import pyautogui as pg
import base64
import re

class PDDOrderRemark:

    def __init__(self, file):
        self.file = file
        self.cookie_dir = os.getcwd() + '\\cookie\\'
        self.data = []
        self.color_list = ["红色", "黄色", "绿色", "蓝色", "紫色"]
        self.export_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) + '.xlsx'

    def get_size(self):
        width, height = pg.size()  # 获取屏幕尺寸
        return width, height

    def wait(self, page, second=1):
        """
        等待
        :param page:
        :return:
        """
        page.wait_for_load_state(state="networkidle")
        time.sleep(second)

    def remove_advertising(self, page):
        """
        删除广告
        :param page:
        :return:
        """
        if page.is_visible('//div[contains(@class,"MDL_footer_")]'):
            page.click('//span[text()="已安装去使用"]')
        if page.is_visible('//div[@class="modal-system-notice-modal-title"]'):
            page.click('//span[text()="关闭"]')
        if page.is_visible('//div[contains(@class, "ImportantList_msgbox-header")]'):
            page.click('//i[contains(@class, "ImportantList_close__")]')

    def login(self, playwright):
        if os.path.exists(self.cookie_dir) is False:
            os.mkdir(self.cookie_dir)

        logger.info("开始登录拼多多后台")
        width, height = self.get_size()  # 获取屏幕尺寸
        executable_path = r"C:\Users\Administrator\Desktop\worker11\ms-playwright\firefox-1234\firefox\firefox.exe"
        browser = playwright.firefox.launch(headless=False,
                                            executable_path=executable_path
                                            )  # 打开火狐浏览器
        self.context = browser.new_context(viewport={"width": width, "height": height - 180})
        self.page = self.context.new_page()
        self.page.goto('https://mms.pinduoduo.com/home/')
        self.remove_advertising(self.page)  # 定义函数，关闭广告
        for t in range(300, 0, -1):
            time.sleep(1)
            if self.page.is_visible('//ul[@class="nav-item-group-content"]/li/a/span[contains(text(),"发布新商品")]'):
                logger.info(' - 登陆成功')
                return True
            if t == 300:
                b_img = self.page.wait_for_selector('//div[@class="qr-code"]').screenshot()  # 截屏
                img_base64 = str(base64.b64encode(b_img))  # base64加密
                img_base64 = re.search("b'(.+)'", img_base64).group(1)  # 按照规则"b'(.+)'"来匹配到的第二个字符串
                logger.debug(f'data:image/png;base64,{img_base64}')
                img = f'<img class="qrcode-img" src="data:image/png;base64,{img_base64}" style="width: 240px;height: 240px;">'
                logger.info(' - 请打开拼多多商家版APP,使用扫码登录。。。')
                logger.info(img)

            if self.page.is_visible('//p[text()="二维码已失效"]'):
                self.page.click('//button/span[text()="点击刷新"]')
                self.wait(self.page)
                b_img = self.page.wait_for_selector('//div[@class="qr-code"]').screenshot()  # 截屏
                img_base64 = str(base64.b64encode(b_img))  # base64加密
                img_base64 = re.search("b'(.+)'", img_base64).group(1)  # 按照规则"b'(.+)'"来匹配到的第二个字符串
                logger.debug(f'data:image/png;base64,{img_base64}')
                img = f'<img class="qrcode-img" src="data:image/png;base64,{img_base64}" style="width: 240px;height: 240px;">'
                logger.info(' - 请打开拼多多商家版APP,使用扫码登录。。。')
                logger.info(img)

            if t % 20 == 0:
                b_img = self.page.wait_for_selector('//div[@class="qr-code"]').screenshot()  # 截屏
                img_base64 = str(base64.b64encode(b_img))  # base64加密
                img_base64 = re.search("b'(.+)'", img_base64).group(1)  # 按照规则"b'(.+)'"来匹配到的第二个字符串
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

    def append_data(self, order_id, flag_color, remark_content, remark_result):
        """
        添加结果
        :param order_id:
        :param flag_color:
        :param remark_content:
        :param remark_result:
        :return:
        """
        date_items= [str(order_id), flag_color, remark_content, remark_result]
        self.data.append(date_items)

    def append_remark(self, row):  # 对当前一条订单进行处理
        """
        添加备注
        :param order_id:订单编号
        :return:
        """
        logger.info(f'当前正在处理{row["订单编号"]}')
        self.page.fill('//input[@placeholder="请输入完整订单编号"]', str(row["订单编号"]))
        with self.page.expect_response('https://mms.pinduoduo.com/mangkhut/mms/recentOrderList', timeout=5000) as response_info:
            self.page.click('//span[text()="查询"]')
        response = response_info.value
        resp_json = json.loads(response.text())
        order_num = resp_json.get("result").get("pageItems")  # 查询订单号,并筛选
        if len(order_num) == 0:
            msg = "查无此单号"
            self.append_data(row['订单编号'], row["旗帜颜色"], row["备注内容"], msg)
            return
        if len(order_num) > 1:
            msg = "该单号有多条订单"
            self.append_data(row['订单编号'], row["旗帜颜色"], row["备注内容"], msg)
            return
        if order_num[0].get("remark_status") == 0:
            self.page.click('//span[text()="添加备注"]')
        elif order_num[0].get("remark_status") == 1:
            self.page.click('//span[text()="修改备注"]')
        old_text = self.page.query_selector('//textarea[@placeholder="订单备注由商家添加，仅平台客服和商家可见"]').text_content()
        new_text = old_text + '\n' + row["备注内容"]
        self.page.fill('//textarea[@placeholder="订单备注由商家添加，仅平台客服和商家可见"]', new_text)
        flag_color = row["旗帜颜色"]
        if flag_color not in self.color_list:
            msg = f"{flag_color}不在系统旗帜颜色选项中"
            self.append_data(row['订单编号'], row["旗帜颜色"], row["备注内容"], msg)
            return
        color_objs = self.page.query_selector_all('//div[contains(@class, "tag-wraper")]')
        color_obj = [i for i in color_objs if i.text_content()==row['旗帜颜色']][0]
        class_str = color_obj.get_attribute('class')
        if 'active-tag' in class_str:
            logger.info('该旗帜颜色已被选中')

        else:
            self.page.click(f'//span[@class="input-text" and text()="{row["旗帜颜色"]}"]')
        with self.page.expect_response('https://mms.pinduoduo.com/pizza/order/noteTag/add', timeout=5000) as response_info:
            self.page.click('//span[text()="保存"]')
        response = response_info.value
        resp_json = response.json()
        err_msg = resp_json.get("errorMsg")
        if err_msg == "成功":
            self.append_data(row['订单编号'], row["旗帜颜色"], row["备注内容"], "成功")
            return
        self.append_data(row['订单编号'], row["旗帜颜色"], row["备注内容"], err_msg)
        self.page.click('//span[text()="取消"]')

    def process(self):
        try:
            self.page.goto('https://mms.pinduoduo.com/orders/list', timeout=10000)
        except Exception:
            pass
        self.wait(self.page)
        self.remove_advertising(self.page)  # 定义函数，关闭广告
        for index, row in self.df.iterrows():  # 遍历读取的Excel表格
            if pd.isna(row['订单编号']):
                continue
            logger.info(f'当前执行{index + 1}条订单, 剩余{len(self.df) - (index + 1 )}条')
            try:
                self.append_remark(row)
            except Exception:
                logger.info(f'第{index + 1}条处理失败')
                continue

    def create_excel(self):
        """
        生成表格
        :return:
        """
        columns = ["平台订单号", "旗帜颜色", "备注内容", "备注结果"]
        df = pd.DataFrame(data=self.data, columns=columns)
        df.to_excel(self.export_file, index=False)

    def send_email(self):
        """"
        发送邮件
        """
        # recipientAddrs = "zoupengfei@uniner.com"
        recipientAddrs = setting.get("recipientAddrs")
        smtpHost = 'smtp.163.com'
        port = 465
        sendAddr = '18016454917@163.com'
        password = 'EPNZMCUJPPMCCVRJ'
        msg = email.mime.multipart.MIMEMultipart()
        msg['from'] = sendAddr  # 发送人邮箱地址
        msg['to'] = recipientAddrs  # 多个收件人的邮箱应该放在字符串中,用字符分隔, 然后用split()分开,不能放在列表中, 因为要使用encode属性
        msg['subject'] = "拼多多-订单备注"
        content = "拼多多-订单备注"  # 内容
        txt = email.mime.text.MIMEText(content, 'plain', 'utf-8')
        msg.attach(txt)
        logger.info('准备添加附件....')
        part = MIMEApplication(open(self.export_file, 'rb').read())
        file_name = os.path.split(self.export_file)[-1]
        part.add_header('Content-Disposition', 'attachment', filename=file_name)  # 给附件重命名,一般和原文件名一样,改错了可能无法打开.
        msg.attach(part)
        logger.info("附件添加成功")
        smtp = smtplib.SMTP_SSL(smtpHost, port)  # 需要一个安全的连接，用SSL的方式去登录得用SMTP_SSL，之前用的是SMTP（）.端口号465或587
        smtp.login(sendAddr, password)  # 发送方的邮箱，和授权码（不是邮箱登录密码）
        smtp.sendmail(sendAddr, recipientAddrs.split(";"), str(msg))  # 注意, 这里的收件方可以是多个邮箱,用";"分开, 也可以用其他符号
        smtp.quit()
        logger.info('邮件发送成功')

    def upload_file(self):
        """
        上传表格
        :return:
        """
        file_name = self.export_file.split('\\')[-1]
        self.send_email()
        # utility.post_file(file_name)
        # url = utility.get_download_url(file_name)
        # logger.info(f'任务完成, 请下载<a href="{url}" target="_blank">拼多多订单备注</a>，查看格式是否正确')

    def main(self, playwright):
        try:
            self.df = pd.read_excel(self.file)
            self.login(playwright)
            self.process()
            self.create_excel()
            self.upload_file()
        except Exception as e:
            logger.info(e)
        finally:
            self.page.close()


if __name__ == '__main__':
    logger.info("*" * 10 + "开始启动拼多多-订单备注机器人")

    # data_file = "乐言-拼多多订单备注.xlsx"
    with sync_playwright() as playwright:
        pdd_or = PDDOrderRemark(data_file)
        pdd_or.main(playwright)