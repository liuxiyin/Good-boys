from playwright.sync_api import sync_playwright
from loguru import logger
from email.mime.application import MIMEApplication
from Crypto.Hash import MD5
import email.mime.multipart
import email.mime.text
import smtplib
import pandas as pd
import requests
import base64
import re
import uuid
import os
import time
import json
import arrow
import math
import zipfile
import shutil
import urllib
import random

# setting = {
#     'start_time': '2021-12-01',
#     'end_time': '2021-12-10',
#     'acc':'wesens旗舰店:sophia',
#     'pwd':'liusuhua6784'
# }


class Taobao:
    def __init__(self):
        self.download_path = os.getcwd() + '\\data\\导出\\' + str(uuid.uuid4())

        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) + '.xlsx'  # 暂时忽视
        self.list1 = []  # 用来保存查询到的 订单数 的列表
        self.t_num = 0  # 下载次数

        if not os.path.exists(os.getcwd() + '\\data\\'):
            os.mkdir(os.getcwd() + '\\data\\')
        if not os.path.exists(os.getcwd() + '\\data\\导出\\'):
            os.mkdir(os.getcwd() + '\\data\\导出\\')
        else:
            path_1 = os.getcwd() + '\\data\\导出\\'
            list_dir = os.listdir(path_1)
            for i in list_dir:
                file_name = os.path.join(path_1 + i)
                if os.path.isfile(file_name):
                    os.remove(file_name)
                elif os.path.isdir(file_name):
                    shutil.rmtree(file_name)
        if not os.path.exists(self.download_path):
            os.mkdir(self.download_path)

    def save_cookies(self, path="cookies"):
        if os.path.exists("cookies") is False:
            os.mkdir('cookies')
        self.context.storage_state(path=f'cookies/{path}.json')

    def load_cookies(self, path="cookies"):
        if os.path.exists(f'cookies/{path}.json'):
            with open(f'cookies/{path}.json') as f:
                json_str = json.load(f)
                self.context.add_cookies(json_str['cookies'])
            return True
        else:
            return False

    def print_img(self, page):
        try:
            if page.is_visible('//div[@class="qrcode-error"]', timeout=3000):
                page.click('//button[@class="refresh"]')
            page.wait_for_selector('//div[@class="qrcode-img"]').screenshot(path='data/1.png')
            with open('data/1.png', 'rb') as f:
                img_64 = str(base64.b64encode(f.read()))
            img_64 = re.search("b'(.+)'", img_64).group(1)
            img = f""" <img class="qr" src="data:image/png;base64,{img_64}" alt="" style="width: 180px;height: 180px;"> """
            logger.info(img)
            logger.info('请打开APP进行扫码登录')
        except Exception:
            return False

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
            return False

    def login(self, playwright):
        self.chrome = playwright.chromium.connect_over_cdp("http://localhost:9002")
        self.context = self.chrome.new_context(accept_downloads=True)
        self.page = self.context.new_page()
        self.goto_url('https://login.taobao.com/member/login.jhtml', self.page)
        if not setting.get('acc') or not setting.get('pwd'):
            log_fun = "扫码"
            self.page.click('//i[@class="iconfont icon-qrcode"]')
        else:
            log_fun = "账密"
            self.page.fill('//input[@name="fm-login-id"]', setting['acc'])
            self.page.fill('//input[@name="fm-login-password"]', setting['pwd'])
            with self.page.expect_navigation():
                self.page.click('//button[text()="登录"]')
        for i in range(300, 0, -1):
            time.sleep(1)
            try:
                if self.page.query_selector('//div[@id="seller-nav"]'):
                    logger.info('登陆成功')
                    return True
            except Exception:
                pass
            if log_fun == "扫码" and i % 30 == 0:
                self.print_img(self.page)
            # self.click_err_img(self.page)

    def goto_url(self, url, page):
        try:
            page.goto(url, timeout=1000 * 10)
        except:
            return

    def wait(self, page, seconds=1):
        try:
            page.wait_for_load_state(state="networkidle", timeout=10000)
        except:
            pass
        finally:
            time.sleep(seconds)

    def remove_file(self):
        """
        移除旧文件
        :return:
        """
        file_list = os.listdir(self.download_path)
        for i in file_list:
            file_path = self.download_path + i
            os.remove(file_path)

    def get_diff_second(self, time1, time2):
        diff = time2 - time1
        seconds = diff.total_seconds()
        return int(seconds)

    def get_total_num(self):
        with self.page.expect_response('https://trade.taobao.com/trade/itemlist/asyncSold.htm?**') as response_info:
            self.page.click('//button[text()="搜索订单"]')
        time.sleep(2)
        response = response_info.value
        result = response.json()
        total_num = result.get('page').get('totalNumber')  # page字典里的totalNumber字典
        return total_num

    def select_year(self, selector, year_str):
        """
        选择年份
        :param selector:
        :param year_str:
        :return:
        """
        self.page.click(selector)
        for i in range(30):
            time.sleep(1)
            year = self.page.query_selector('//a[@class="rc-calendar-year-select"]').text_content()
            if year_str == year:
                return True
            elif int(re.findall(r"\d+", year_str)[0]) < int(re.findall(r"\d+", year)[0]):
                self.page.click('//a[@title="上一年 (Control键加左方向键)"]')
            else:
                self.page.click('//a[@title="下一年 (Control键加右方向键)"]')
        return False

    def select_month(self, selector, month_str):
        """
        选择月份
        :param selector:
        :param month_str:
        :return:
        """
        if self.page.is_visible('//a[@class="rc-calendar-month-select"]') is False:
            self.page.click(selector)
        for i in range(30):
            month = self.page.query_selector('//a[@class="rc-calendar-month-select"]').text_content()
            if month_str == month:
                return True
            elif int(re.findall(r"\d+", month_str)[0]) < int(re.findall(r"\d+", month)[0]):
                self.page.click('//a[@title="上个月 (翻页上键)"]')
            else:
                self.page.click('//a[@title="下个月 (翻页下键)"]')
        return False

    def select_date(self, date_str):
        """
        选择日期
        :param date_str:
        :return:
        """
        self.page.click(f'//td[@title="{date_str}"]')

    def select_hour(self, hour=0):
        self.page.click('//input[@class="rc-calendar-time-input"]')
        self.page.click(f'//table[@class="rc-calendar-time-panel-table"]/tbody/tr/td/a[text()="{hour}"]')

    def select_minute(self, minute=0):
        self.page.click('//span[@class="rc-calendar-time-minute"]/input[@class="rc-calendar-time-input"]')
        self.page.click(f'//table[@class="rc-calendar-time-panel-table"]/tbody/tr/td/a[text()="{minute}"]')

    def select_second(self, second=0):
        self.page.click('//span[@class="rc-calendar-time-second"]/input[@class="rc-calendar-time-input"]')
        self.page.click(f'//table[@class="rc-calendar-time-panel-table"]/tbody/tr/td/a[text()="{second}"]')

    def select_time(self, selecor, time_obj):
        """
        选择时间
        :param selecor:
        :param time_obj:
        :return:
        """
        self.select_year(selector=selecor, year_str=str(time_obj.year) + '年')
        self.select_month(selector=selecor, month_str=str(time_obj.month) + '月')
        self.select_date(date_str=time_obj.format('YYYY-M-D'))

    def download_file(self):
        """
        下载文件
        :return:
        """
        self.t_num = self.t_num + 1
        self.list1[0] = self.list1[0] - self.list1[-1]
        if self.page.query_selector('//div[@class="batch-exports-mod__container___2oh8X"]'
                                    '').get_attribute("style") == "":
            pass
        else:
            self.page.click('//button[text()="批量导出"]')
        self.page.click('//button[text()="生成报表"]')
        with self.page.expect_popup() as popup_info:
            self.page.click('//button[text()="确定"]')
        new_page = popup_info.value
        self.wait(new_page)
        t1 = arrow.get(arrow.now()).shift(seconds=60 * 5)
        for i in range(360, 0, -1):
            try:
                if new_page.is_visible('//ul/li[1]/div/a[@title="下载订单报表"]', timeout=100000):
                    break
                time.sleep(2)
                if i % 30 == 0:
                    logger.info(f'正在下载第{self.t_num}份报表，预计还剩{self.list1[0]}条订单')
            except:
                pass
        for i in range(10):
            with new_page.expect_download(timeout=1000 * 60 * 5) as download_info:
                new_page.query_selector('//ul/li[1]/div/a[@title="下载订单报表"]').click()
            download = download_info.value
            if download:
                break
            time.sleep(1)
        str_name = download.suggested_filename
        file_path = self.download_path + download.suggested_filename
        download.save_as(file_path)
        self.save_file(file_path, str_name)
        # self.updata_file(strs=file_path)  # 下载
        new_page.on('dialog', lambda dialog: dialog.accept())  # 检测到突然弹出的一个框
        new_page.click('//ul/li[1]/div/a[@title="发送密码"]')
        new_page.close()
        logger.info('报表生成完毕, 请稍后...')
        if arrow.now() < t1 and self.list1[0] != 0:
            wait_second = self.get_diff_second(arrow.now(), t1)
            logger.info(f'请等待{wait_second}秒,之后开始下一次下载,预计还剩{self.list1[0]}条订单')
            time.sleep(wait_second)

    def save_file(self, file_path, str_name):
        # new_file_path = os.path.join(self.download_path,"result_"+ uuid.uuid4().hex + os.path.splitext(file_path)[-1])
        new_file_path = os.path.join(self.download_path, "result_" + str_name)
        with open(file_path, 'rb') as f_rb:
            with open(new_file_path, 'wb') as f_wb:
                f_wb.write(f_rb.read())

    def zip_dir(self, folder_path):
        if not os.path.isdir(folder_path):
            print('%s不是一个文件夹' % folder_path)
            return

        shop_name = setting.get('acc').split(":")[0]
        file_name1 = shop_name + "_" + setting.get("start_time") + "_" + setting.get("end_time") + "天猫订单"
        zip_file_path = os.getcwd() + '\\data\\导出\\' + file_name1 + '.zip'

        # zip_file_path = folder_path + '.zip'
        z = zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED)
        for dir_path, dir_names, file_names in os.walk(folder_path):
            f_path = dir_path.replace(folder_path, '')
            f_path = f_path and f_path + os.sep or ''
            for file_name in file_names:
                z.write(os.path.join(dir_path, file_name), f_path + file_name)
                print('压缩成功')
        z.close()
        logger.info(f'压缩文件所在地址是:{zip_file_path}')

        return zip_file_path

    def process(self):
        self.goto_url(url='https://trade.taobao.com/trade/itemlist/list_sold_items.htm', page=self.page)
        time.sleep(5)
        s_time = arrow.get(setting['start_time'])  # 开头时间
        e_time = arrow.get(setting['end_time'])
        e_time = arrow.get(e_time.year, e_time.month, e_time.day, 23, 59, 59)  # 结尾时间
        e_time1 = arrow.get(e_time.year, e_time.month, e_time.day, 23, 59, 59)  # 中间时间

        little_num = 8000  # 最小数目-------------------------
        big_num = 10000  # 最大数目-----------------------  //button[text()="显示更多页码"]
        self.page.wait_for_selector(selector='//input[@placeholder="请选择时间范围起始"]', timeout=1000*60*5)
        if self.page.is_visible('//button[text()="显示更多页码"]'):
            self.page.click('//button[text()="显示更多页码"]')
            time.sleep(2)
        while True:
            self.select_time(selecor='//input[@placeholder="请选择时间范围起始"]', time_obj=s_time)
            self.select_hour(hour=s_time.hour)
            self.select_minute(minute=s_time.minute)
            self.select_second(second=s_time.second)
            self.page.click('//span[@class="rc-calendar-footer-btn"]/a[text()="确定"]')
            self.select_time(selecor='//input[@placeholder="请选择时间范围结束"]', time_obj=e_time1)
            self.select_hour(hour=e_time1.hour)
            self.select_minute(minute=e_time1.minute)
            self.select_second(second=e_time1.second)
            self.page.click('//span[@class="rc-calendar-footer-btn"]/a[text()="确定"]')
            time.sleep(1)
            for i in range(20):
                try:
                    total_num = self.get_total_num()
                    break
                except:
                    time.sleep(30)
            logger.info(total_num)
            self.list1.append(total_num)
            if total_num == 0 and e_time1 == e_time:
                logger.info("该时间段内无订单")
                return False

            if total_num < little_num and e_time1 != e_time:
                low = self.get_diff_second(s_time, e_time1)  # 最小时间段
                high = self.get_diff_second(s_time, e_time)  # 最大时间段 > 最小时间段
                i = 0
                while low < high and i < 50:
                    i = i + 1
                    seconds = math.ceil((low + high) / 2)
                    e_time1 = arrow.get(s_time).shift(seconds=seconds)
                    self.select_time(selecor='//input[@placeholder="请选择时间范围结束"]', time_obj=e_time1)
                    self.select_hour(hour=e_time1.hour)
                    self.select_minute(minute=e_time1.minute)
                    self.select_second(second=e_time1.second)
                    self.page.click('//span[@class="rc-calendar-footer-btn"]/a[text()="确定"]')
                    time.sleep(1)
                    total_num = self.get_total_num()
                    logger.info(total_num)
                    if total_num is None:
                        total_num = 0
                    if total_num < little_num:
                        low = self.get_diff_second(s_time, e_time1)  # e_time2作为low
                    elif total_num > big_num:
                        high = self.get_diff_second(s_time, e_time1)  # e_time2作为high
                    else:
                        break

            if total_num < big_num and e_time1 == e_time:
                logger.info(f'本次下载的开始时间:{s_time.format("YYYY-MM-DD HH:mm:ss")}')
                logger.info(f'本次下载的结束时间:{e_time1.format("YYYY-MM-DD HH:mm:ss")}')
                logger.info(f'本次下载订单总量{total_num}条')
                self.list1.append(total_num)
                self.download_file()
                break
            elif total_num < big_num and e_time != e_time1:
                logger.info(f'本次下载的开始时间:{s_time.format("YYYY-MM-DD HH:mm:ss")}')
                logger.info(f'本次下载的结束时间:{e_time1.format("YYYY-MM-DD HH:mm:ss")}')
                logger.info(f'本次下载订单总量{total_num}条')
                self.list1.append(total_num)
                s_time = arrow.get(e_time1).shift(seconds=1)
                e_time1 = e_time
                self.download_file()
                # for i in range(5, 0, -1):
                #     logger.info(f'等待{i}分钟后进行下一次下载')
                #     time.sleep(60)
                logger.info('开始下载')
                continue
            seconds = self.get_diff_second(s_time, e_time1) // 2  # 求开始时间到查询用结束时间 差多长时间，取一半
            if seconds == 0:
                logger.info('开始时间和结束时间是同一个时间，无法下载')
                return False
            e_time1 = arrow.get(s_time).shift(seconds=seconds)  # 数目太多，开始时间到查询用结束时间 差距只用一半
        return True

    def create_file(self):
        """
        合并文件，生成输出报表
        :return:
        """
        data = []
        file_list = os.listdir(self.download_path)
        for i in file_list:
            file_path = self.download_path + i
            df = pd.read_excel(file_path)
            data.append(df)
        df_all = pd.concat(data).reset_index(drop=True)
        df_all.to_excel(self.result_file, index=False)

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
        msg['subject'] = "千牛文件"
        content = "千牛文件"  # 内容
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
        from loguru import logger
        # robotNo = serial_no
        # robotPwd = robot_secret

        obj = MD5.new()
        obj.update(robot_secret.encode("utf-8"))  # gb2312 Or utf-8
        robotPwd = obj.hexdigest()

        url = f'{base_url}/worker/internal/file/upload/result'
        # url = f'{base_url}/file/upload/{robotNo}/{robotPwd}'
        header = {
            'robotNo': serial_no,
            'robotPwd': robotPwd

        }
        data = {}
        # 1为文件类型

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
        time.sleep(10)
        return
        s_time = arrow.get(setting['start_time'])  # 开头时间
        e_time = arrow.get(setting['end_time'])  # 结束时间
        e_time = arrow.get(e_time.year, e_time.month, e_time.day, 23, 59, 59)  # 结束时间
        seconds = self.get_diff_second(s_time, e_time)
        if s_time > e_time:
            logger.info("请确保[开始时间]小于[结束时间]，并重新启动机器人")
            return
        elif s_time > arrow.now():
            logger.info("请确保[开始时间]不超过[当日时间]，并重新启动机器人")
            return
        elif e_time > arrow.now():
            logger.info("请确保[结束时间]不超过[当日时间]，并重新启动机器人")
            return
        elif seconds > 60 * 60 * 24 * 31 or seconds == 0:
            logger.info("请确保[开始时间]和[结束时间]的间隔不为0且小于31天，并重新启动机器人")
            return
        self.login(playwright)
        self.process()
        logger.info("下载结束")
        zip_file_path = self.zip_dir(self.download_path)
        try:
            self.updata_file(zip_file_path)
        except:
            pass
        return True


if __name__ == '__main__':
    try:
        import subprocess

        server = subprocess.Popen(
            r"./ms-playwright/chromium-907428/chrome-win/chrome --remote-debugging-port=9002 ")
    except Exception as e:
        os.popen(
            './chrome --no-sandbox --disable-gpu --disable-dev-shm-usage --use-gl=desktop --window-size=1600,1024 --remote-debugging-port=9002')
        time.sleep(2)
    with sync_playwright() as playwright:
        try:
            tb = Taobao()
            tb.main(playwright)
        except Exception as e:
            logger.info(repr(e))
        finally:
            server.kill()
