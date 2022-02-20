from playwright.sync_api import sync_playwright
from email.mime.application import MIMEApplication
from loguru import logger
from Crypto.Hash import MD5
import requests
import os
import uuid
import base64
import re
import pandas as pd
import json
import email.mime.multipart
import email.mime.text
import smtplib
import time
import shutil
import arrow
import zipfile
import csv

#
# setting = {
#     'url': "https://www.erp321.com/",
#     'account': "18626197197",
#     'password': "BN20210407@ "
# }


# 聚水潭
downloads_path = os.getcwd() + '\\data\\财务对账\\'
executable_path = "./ms-playwright/firefox-1281/firefox/firefox.exe"
executable_path1 = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
executable_path1 = executable_path1 if os.path.exists(executable_path1) else None
#
# data = {
#     "account": "loveebuyer",
#     "pwd": "Zpf19870501..",
#     "x": "BN20210407@ ",
#     "y": "18626197197"
#
# }

# setting = {
#     "bill_type": "月账单"
# }


class FinancialReconciliation:
    """
    财务对账
    """
    def __init__(self):
        self.result_file = os.getcwd() + '\\data\\' + str(uuid.uuid4()) + '.xlsx'
        self.data = []

    def get_file_ju(self,playwright):
        """
        下载聚水潭文件
        :param playwright:
        :return:
        """
        ju_shui_tan = JuShuiTan()
        file_path = ju_shui_tan.main(playwright)
        return file_path
        # return r"D:\GITFile\发货异常\data\聚水潭-销售主体分析.xlsx"

    def get_file_ali_pay(self, playwright):
        """
        下载支付宝账单
        :param playwright:
        :return:
        """
        ali_pay = Alipay()
        file_path = ali_pay.main(playwright)
        return file_path

    def get_msg_ali(self, row):
        """
        获取阿里表格数据
        :param row:
        :return:
        """
        data_dict = {
            "是否正常": "",
            "收款金额": "",
        }
        if pd.isna(row['线上订单号']):
            return False
        id1 = 'T200P' + row['线上订单号']
        pay_money = row['已付金额']
        df = self.df_ali[self.df_ali['商户订单号'] == id1]
        if len(df) == 0:
            data_dict['是否正常'] = '未找到对应的单号'
            return data_dict
        if len(df) > 1:
            data_dict['是否正常'] = '找到多条单号'
            return data_dict
        money = df['收入金额']
        if pay_money == money:
            data_dict['是否正常'] = '正常'
            data_dict['收款金额'] = money
            return data_dict
        data_dict['是否正常'] = '异常'
        return data_dict

    def append_data(self, row, data_dict):
        """
        添加数据
        :param row:
        :param data_dict:
        :return:
        """
        col = ['店铺', '订单状态', '原始线上订单号', '商品编码', '款式编码', '名称',
               '运费收入', '买家账号', '订单日期', '发货日期', '付款日期', '确认收货日期',
               '订单类型', '发货仓', '快递公司', '快递单号', '订单重量', '产品分类', '成本价',
               '虚拟分类', '颜色规格', '供应商', '售后确认日期', '售后分类', '问题类型',
               '销售成本', '实发成本', '应付金额', '当期退货成本', '当期实退成本', '售价',
               '优惠金额', '当期退货数量', '线上颜色规格']

        data = {}
        for i in col:
            data[i] = row.get(i)
        data['是否异常'] = data_dict.get("是否正常")
        data['收款金额'] = data_dict.get("收款金额")
        self.data.append(data)

    def table_comparison(self):
        """
        表格对比
        :return:
        """
        for index, row in self.df_jst.iterrows():
            logger.info(f'当前正在执行{index +1}条')
            data_dict = self.get_msg_ali(row)
            if data_dict is False:
                continue
            self.append_data(row, data_dict)
        logger.info('表格比较完成')

    def send_email(self):
        """"
        发送邮件
        """
        # recipientAddrs = setting["recipientAddrs"]
        recipientAddrs = setting.get("recipientAddrs")
        smtpHost = 'smtp.163.com'
        port = 465
        sendAddr = '18016454917@163.com'
        password = 'EPNZMCUJPPMCCVRJ'
        msg = email.mime.multipart.MIMEMultipart()
        msg['from'] = sendAddr  # 发送人邮箱地址
        msg['to'] = recipientAddrs  # 多个收件人的邮箱应该放在字符串中,用字符分隔, 然后用split()分开,不能放在列表中, 因为要使用encode属性
        content = "财务对账报表·"  # 内容
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

    def get_df_ali(self,file):
        """
        读取csv
        :param file:
        :return:
        """
        with open(file, 'r') as f:
            # reader = csv.reader(f)
            # result = list(reader)
            data = f.read().replace('\t', '')
        result = data.split('\n')
        f.close()
        index_list = [index for index, val in enumerate(result) if "-账务明细列表-" in val or "-账务明细列表结束-" in val]  # 表的读取范围
        data_list = result[index_list[0]+1: index_list[-1]]  # 按照范围读取表的内容
        if len(data_list) == 0:
            return False
        columns = data_list[0].split(',')  # 获取列名
        data_list_1 = data_list[1:]
        data_val = [i.split(',') for i in data_list_1]  # 获取内容
        df = pd.DataFrame(columns=columns, data=data_val)  # 将列名，内容 按照dataframe格式排列好
        return df

    def data_sum(self):
        """
        数据汇总
        :return:
        """

    def updata_file(self, strs, nass_name='文件'):
        obj = MD5.new()
        obj.update(robot_secret.encode("utf-8"))  # gb2312 Or utf-8
        robotPwd = obj.hexdigest()  # 十六进制加密

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

    def process_excel(self, jst_file, ali_file):
        """
        表格处理
        :param jst_file:
        :param ali_file:
        :return:
        """
        self.df_jst = pd.read_excel(jst_file, dtype={"线上订单号": "str"})
        self.df_ali = self.get_df_ali(ali_file)
        if self.df_ali is False:
            logger.info('支付宝账单无可用数据')
            return
        self.table_comparison()  # 表格对比
        self.data_sum()  # 汇总
        df = pd.DataFrame(data=self.data)
        df.to_excel(self.result_file, index=False)
        if setting.get("recipientAddrs"):
            self.send_email()
        self.updata_file(strs=self.result_file)

    def main(self,playwright):
        if not setting.get("bill_type"):
            logger.info('支付宝账单类型未填写，请填写后重新运行机器人')
            return
        jst_file = self.get_file_ju(playwright)
        if jst_file is False:
            logger.info('聚水潭无文件可下载')
            return
        logger.info('聚水潭文件下载完毕')
        ali_file = self.get_file_ali_pay(playwright)
        if ali_file is False:
            logger.info('支付宝账单无文件可下载')
            return
        logger.info('支付宝账单下载完毕')
        self.process_excel(jst_file, ali_file)


class JuShuiTan:
    def __init__(self):
        pass

    def move_olf_file(self):
        """
        移除历史文件
        :return:
        """
        old_file_dir = os.getcwd() + '\\历史文件\\'
        if os.path.exists(old_file_dir) is False:
            os.mkdir(old_file_dir)
        file_list = os.listdir(downloads_path)
        for file in file_list:
            file_path = downloads_path + file
            if os.path.isdir(file_path):
                continue
            file_path_d = old_file_dir + file
            shutil.move(file_path, file_path_d)
        logger.info('历史文件迁移完毕')

    def print_img(self, page):
        img_64 = page.wait_for_selector('//div[@id="qrcode"]/img').get_attribute('src')
        img = f""" <img class="qr" src="{img_64}" alt="" style="width: 180px;height: 180px;">
                            """
        logger.info(' - 请打开钉钉扫码登录')
        logger.info(img)

    def refreash(self, page):
        try:
            if page.is_visible('//div[@class="login_qrcode_refresh"]'):
                page.click('//span[@id="refreashQrCodeBtn"]')
        except Exception:
            return

    def add_cookie(self, cookie_path):
        if os.path.exists(cookie_path):
            try:
                f = open(cookie_path, 'r', encoding = "utf-8")
                read_f = json.loads(f.read())
                f.close()
                self.context.add_cookies(read_f.get("cookie"))
            except Exception:
                return

    def login(self, playwright):

        try:
            self.browser = playwright.chromium.launch(headless=False,
                                                     executable_path=executable_path1,
                                                     downloads_path=downloads_path
                                                     )
            self.context = self.browser.new_context(accept_downloads=True)
            self.page = self.context.new_page()
            self.page.goto('https://www.erp321.com/login.aspx')
            self.page.wait_for_load_state("networkidle")
            self.page.click('//div[@class="login-switch"]')
            self.page.click('//div[text()="钉钉扫码登录"]')
            time.sleep(1)
            for i in range(300, 0, -1):
                try:
                    if self.page.is_enabled('//div[@id="main"]', timeout=2000):
                        logger.info('聚水潭登录完毕')
                        return True
                except Exception:
                    pass
                if i % 30 == 0:
                    logger.info('聚水潭还未登录，请扫码登录')
                    login_iframe = self.page.wait_for_selector('//div[@id="login_container"]/iframe').content_frame()
                    self.print_img(login_iframe)
                self.refreash(login_iframe)
                time.sleep(1)
            return False
        except Exception as e:
            print(f'error：{e}')
            return False

    def remove_alert(self):
        """
        去除弹窗
        :return:
        """
        for i in range(60):
            if  self.frame.wait_for_selector('//iframe[@id="float_frame"]'):
                break
            time.sleep(1)
            continue
        frame = self.frame.wait_for_selector('//iframe[@id="float_frame"]').content_frame()
        if frame.is_visible('//span[@title="关闭"]'):
            frame.click('//span[@title="关闭"]')

    def select_date(self, frame_filter):
        """
        选择时间
        :return:
        """
        frame_filter.wait_for_selector('//span[@id="date_desc"]').click()
        time.sleep(2)
        frame_filter.click('//div[text()="搜确认收货时间"]')
        if setting['bill_type'] == "日账单":
            frame_filter.click('//div[text()="昨天"]')
            frame_filter.click('//input[@title="内部订单号"]')
            return
        if setting["bill_type"] == "月账单":
            last_month = arrow.now().shift(months=-1)
            this_month = arrow.get(arrow.now().year, arrow.now().month, 1)
            start_time = arrow.get(last_month.year, last_month.month, 1).format('YYYY-MM-DD')
            end_time = this_month.shift(days=-1).format('YYYY-MM-DD')
            frame_filter.click('//input[@title="内部订单号"]')
            frame_filter.click('//input[@id="order_date_begin"]')
            frame_filter.fill('//input[@id="order_date_begin"]', start_time)
            frame_filter.press('//input[@id="order_date_begin"]', 'Enter')
            frame_filter.click('//input[@id="order_date_end"]')
            frame_filter.fill('//input[@id="order_date_end"]', end_time)
            frame_filter.press('//input[@id="order_date_end"]', 'Enter')

    def update_file(self):
        """
        修改文件后缀
        :return:
        """
        logger.info('开始下载表格文件')
        for i in range(360):
            file_list = os.listdir(downloads_path)
            if len(file_list) == 0:
                time.sleep(1)
                continue

        for i in range(360):
            try:
                file_list = os.listdir(downloads_path)
                file = file_list[0]
                if file.split(".")[1] == "crdownload":
                    time.sleep(2)
                    logger.info('正在下载')
                    continue
            except:
                break
        logger.debug('开始转换')
        old_name = downloads_path + file
        file_name = downloads_path + '聚水潭-销售主体分析.xlsx'
        os.rename(old_name, file_name)
        return file_name

    def process(self):
        """
        业务逻辑处理
        :return:
        """
        self.frame = self.page.wait_for_selector('//div[@id="frame_list"]/iframe[@tagname="home"]').content_frame()
        self.remove_alert()
        self.page.click('//div[@id="side"]/div[@class="side-menu"]/div[text()="报表"]')  # 选择菜单-报表
        self.page.click('//div[contains(text(), "销售主题分析(财务)")]')
        frame_finance = self.page.wait_for_selector('//iframe[contains(@src, "finance")]').content_frame()  # 财务iframe
        frame_filter = frame_finance.wait_for_selector('//iframe[@id="s_filter_frame"]').content_frame()
        time.sleep(5)
        # 选择到货时间确认
        self.select_date(frame_filter)
        frame_filter.click('//div[@id="reload_rpt"]/span/span[text()="生成报表"]')
        frame_finance.wait_for_selector('//div[text()="明细(订单商品)"]').click()
        frame_channel = frame_finance.wait_for_selector('//iframe[@data-url="detail.aspx"]').content_frame()
        frame_channel.wait_for_load_state(state="networkidle")
        try:
            frame_channel.wait_for_selector('//span[@title="导出当前表格数据"]').click()
        except Exception:
            logger.info('未找到导出表格')
            return False
        file_path = self.update_file()
        return file_path

    def main(self, playwright):
        if os.path.exists(downloads_path) is False:
            os.mkdir(downloads_path)
        self.move_olf_file()
        log_flg = self.login(playwright)
        if log_flg is False:
            logger.info('登录失败,请重新启动机器人')
            return False
        file_path = self.process()
        self.browser.close()
        return file_path


class Alipay:
    def __init__(self):
        self.file_list = []
        self.cookie_file = os.getcwd() + '\\cookies\\'
        self.month_dict = { 1:'一', 2:'二',3:'三', 4:'四',5:'五', 6:'六',7:'七', 8:'八',9:'九', 10:"十", 11: "十一", 12:"十二"}

    def print_image(self):
        self.page.query_selector('//div[@id="J-barcode-container"]').screenshot(path="data/log.png")
        with open('data/log.png', 'rb') as f:
            img_64 = str(base64.b64encode(f.read()))
        img_64 = re.search("b'(.+)'", img_64).group(1)
        img = f""" <img class="qr" src="data:image/png;base64,{img_64}" alt="" style="width: 180px;height: 180px;"> """
        logger.info(img)
        logger.info('请打开支付宝APP进行扫码登录')

    def wait(self):
        """
        :return:
        """
        try:
            self.page.wait_for_load_state(state="networkidle", timeout=10000)
        except:
            return

    def login(self, playwright):
        """
        登录
        :return:
        """
        self.down_dir = os.getcwd() + '\\data\\'
        browser = playwright.chromium.launch(headless=False,
                                             executable_path=executable_path1,
                                             downloads_path=downloads_path)
        self.context = browser.new_context(accept_downloads=True)
        self.page = self.context.new_page()
        self.go_to_page(url='https://auth.alipay.com/login/index.htm')
        self.page.click('//li[text()="扫码登录"]')
        for i in range(300, 0, -1):
            try:
                time.sleep(1)
                if i % 60 == 0:
                    self.print_image()
                if self.page.is_visible('//div[@class="qrcode-error"]'):
                    self.page.click('//button[@class="refresh"]')
                if self.page.wait_for_selector('//a[text()="对账中心"]', timeout=60000):
                    logger.info('支付宝登录成功')
                    return True
            except Exception as e:
                logger.info(e)
                browser.close()
                return False

    def down_file_day(self, td_obj):
        """
        下载文件
        :return:
        """
        with self.page.expect_download() as download_info:
            td_obj.query_selector('//i[@aria-label="图标: download"]').click()
        download = download_info.value
        file_name = download.suggested_filename
        file_path = os.path.join(self.down_dir, file_name)
        download.save_as(file_path)
        logger.info('支付宝文件下载完毕')
        return file_path

    def go_to_page(self, url):
        try:
            with self.page.expect_navigation(timeout=10000):
                self.page.goto(url)
        except Exception:
            pass

    def click_item(self):
        for i in range(10):
            button_obj = self.page.query_selector('//div[contains(@class, "ant-btn-group calendarType__")]/button[2]')
            if "default" in button_obj.get_attribute('class'):
                button_obj.click()
                time.sleep(0.5)
                continue
            return


    def bill_downlaod(self):
        """
        账单下载
        :return:
        """
        self.go_to_page(url='https://mbillexprod.alipay.com/enterprise/mbillDownload.htm#/fundBill')
        logger.info('正在进入支付宝对账中心')
        self.page.wait_for_selector('//span[text()="账单下载"]')
        time.sleep(1)
        if setting.get("bill_type") == "日账单":
            yesterday = arrow.now().shift(days=-2).format('YYYY年M月D日')
            td_obj = self.page.query_selector(f'//tr/td[@title="{yesterday}"]')
            if td_obj.query_selector('//i[@aria-label="图标: download"]'):
                file_path = self.down_file_day(td_obj)
                return file_path
            return False
        elif setting.get("bill_type") == "月账单":
            self.wait()
            self.click_item()
            last_month = arrow.now().shift(months=-1).format('YYYY-MM')
            month = arrow.get(last_month).month
            year = arrow.get(last_month).year
            if year != arrow.now().year:
                self.page.click('//div[contains(@class, "leftArr___")]')
                time.sleep(1)
            td_obj = self.page.wait_for_selector(f'//tr/td[@title="{self.month_dict.get(month)}月"]')
            if td_obj.query_selector('//div/div/div/i[@aria-label="图标: download"]'):
                div_obj = td_obj.query_selector('//div/div/div[contains(@class, "downloadIcon___")]')
                file_path = self.down_file_day(div_obj)
                return file_path
            return False
        else:
            return False

    def get_excel(self,file):
        """
        解压文件
        :return:
        """
        zip_file = zipfile.ZipFile(file)
        for file_name in zip_file.namelist():
            try:
                file1 = file_name.encode('cp437').decode('GBK')
            except:
                file1 = file_name.encode('utf-8').decode('utf-8')
            if '汇总' not in file1:
                zip_file.extract(file_name, path=os.getcwd() + '\\data\\')  # 把file_name解压到path
                shutil.move(f'data/{file_name}', f'data/{file1}')  # 文件移动
                excel_path = os.getcwd() + '\\data\\' + file1
                return excel_path
        return False

    def process(self):
        """
        处理流程
        :return:
        """
        file_path = self.bill_downlaod()  # 账单下载
        if file_path:
            excel_file = self.get_excel(file_path)
            return excel_file
        return False

    def main(self, playwright):
        self.login(playwright)
        file_path = self.process()
        self.page.close()
        return file_path


if __name__ == '__main__':
    with sync_playwright() as playwright:
        try:
            ju_shui_tan = FinancialReconciliation()
            ju_shui_tan.main(playwright)
        except Exception as e:
            logger.info(e)
        # try:
        #     al = Alipay()
        #     al.main(playwright)
        # except Exception as e:
        #     print(e)