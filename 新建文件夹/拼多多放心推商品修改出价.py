import gc
import json
import os
import re
import time
import uuid
import pandas as pd
import psutil
import requests
from Crypto.Hash import MD5
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font
from playwright.sync_api import sync_playwright
from loguru import logger
import datetime
import arrow
import shutil
import subprocess
import base64


@staticmethod
def kill_chrome():
    logger.info(" - 清理多余任务")
    pids = psutil.pids()
    for pid in pids:
        try:
            p = psutil.Process(pid)
            process_name = p.name()
            if "chrome" in process_name:
                os.system("kill -9 %s" % pid)
        except Exception as e:
            pass
    gc.collect()
    logger.info(" - 清理多余完毕")

def open_url(page, url):
    """
    打开网址
    :param page:
    :param url:
    :return:
    """
    try:
        page.goto(url, timeout=10000)
    except:
        pass


def print_img(img_path):
    """
    打印二维码
    :param img_path:
    :return:
    """
    with open(img_path, 'rb') as f:
        img_64 = str(base64.b64encode(f.read()))
    img_64 = re.search("b'(.+)'", img_64).group(1)
    img = f""" <img class="qr" src="data:image/png;base64,{img_64}" alt="" style="width: 180px;height: 180px;"> """
    logger.info(img)
    logger.info('请扫码登录')


def get_code():
    time.sleep(2)
    t = time.time()
    t_time = int(round(t * 1000))
    n = 0
    a = {
        "identity": "41E37028-D825-6F22-5F07-3219510DF96B",
        "title": "全类型测试表单",
        "timeout": 3000,
        "createtime": str(f"{t_time}"),
        "details": [
            {
                "type": "1",
                "cname": "手机验证码",
                "label": "phone_code",
                "description": "请输入手机验证码",
                "value": "",
                "rule": "true",
                "imgUrl": "",
                "options": []
            },
            {
                "type": "0",
                "cname": "提交1",
                "label": "subm",
                "description": "提交吧，累了",
                "value": "",
                "rule": "true",
                "imgUrl": "",
                "options": []
            }
        ]
    }
    content = None
    msg_uid = post_msg_center(a)
    while n <= 30:
        n += 1
        time.sleep(10)
        resp = get_msg_center(msg_uid)
        if resp.get("code") == 0:
            content = json.loads(resp.get("data"))
            break
    logger.info(f"验证码为{content['details'][0]['value']}")
    return content['details'][0]['value']


def wait_page(page, seconds=1):
    """
    等待页面加载完成
    :param page:
    :param seconds:
    :return:
    """
    try:
        page.wait_for_load_state(state="networkidle", timeout=5000)
    except:
        pass
    finally:
        time.sleep(seconds)


class PDD:
    """
    拼多多登录
    """
    def __init__(self, page, account=None, pwd=None):
        self.pdd_page = page
        self.account = account
        self.pwd = pwd
        self.url = 'https://mms.pinduoduo.com/goods/goods_list'
        self.data_dict = {'商品创建时间': [], '商品ID': [], '出价金额': [], '修改结果': []}
        self.path = os.getcwd()+'\\data\\修改金额'+'.xlsx'

        if os.path.exists(os.getcwd() + '\\data\\') is False:
            os.mkdir(os.getcwd() + '\\data\\')
        else:
            path_1 = os.getcwd() + '\\data\\'
            list_dir = os.listdir(path_1)
            for i in list_dir:
                file_name = os.path.join(path_1 + i)
                if os.path.isfile(file_name):
                    os.remove(file_name)
                elif os.path.isdir(file_name):
                    shutil.rmtree(file_name)

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

    def login(self):
        open_url(self.pdd_page, self.url)
        wait_page(self.pdd_page, seconds=10)
        for i in range(360):
            try:
                if self.pdd_page.is_visible('//div[@class="user-name"]'):
                    logger.info('拼多多后台登录完成')
                    return True
                if self.account or self.pwd:
                    if i % 30 == 0:
                        self.pdd_page.click('//div[text()="账户登录"]')
                        self.pdd_page.fill('//input[@id="usernameId"]', self.account)
                        self.pdd_page.fill('//input[@id="passwordId"]', self.pwd)
                        self.pdd_page.press('//input[@id="passwordId"]', 'Enter')
                        time.sleep(1)
                else:
                    if i % 30 ==0:
                        self.pdd_page.click('//div[text()="扫码登录"]')
                        img_path = 'data/1.png'
                        self.pdd_page.query_selector('//div[@class="qr-code"]').screenshot(path=img_path)
                        print_img(img_path)
                if self.pdd_page.is_visible('//input[@placeholder="请输入短信验证码"]'):
                    self.pdd_page.click('//a[@data-testid="beast-core-button-link"]')
                    code = get_code()
                    self.pdd_page.fill('//input[@placeholder="请输入短信验证码"]', code)
                    self.pdd_page.click('//button/span[text()="确认"]')
                if i == 180:
                    logger.info('拼多多后台登录超时')
                    return False
            except Exception as e:
                logger.debug(e)
            finally:
                time.sleep(1)

    def get_data(self):
        t1 = int(datetime.datetime.strptime(setting['start_time'], '%Y-%m-%d %H:%M').strftime('%Y%m%d%H%M'))
        t2 = int(datetime.datetime.strptime(setting['end_time'], '%Y-%m-%d %H:%M').strftime('%Y%m%d%H%M'))
        sum = self.pdd_page.text_content('//li[@class="PGT_totalText_1vsdqgl TB_pgtTotalText_1vsdqgl"]')  # 获取货物总数
        pattern = re.compile(r'\d+')
        sum = int(pattern.findall(sum)[0])
        if sum % 10 == 0: num = sum//10
        else: num = sum//10 + 1
        logger.info(f'查询到的在售商品有{sum}个,共{num}页')
        list_n = []
        for i in range(num):
            if i % 10 == 0: logger.info('正在处理')
            with self.pdd_page.expect_response(
                    'https://mms.pinduoduo.com/vodka/v2/mms/query/display/mall/goodsList') as response_info:
                if i == 0:
                    self.pdd_page.click('//span[text()="查询"]')  # 点击查询
                else:
                    self.pdd_page.click('//li[@data-testid="beast-core-pagination-next"]')  # 点击下一页
            time.sleep(2)
            self.wait_for_selector('//tbody/tr[1]')  # 等待货物的列表出现
            response = response_info.value
            result = response.json().get('result')
            goods_list = result.get('goods_list')
            flag = 0
            for j in goods_list:
                list1 = [None, None, None, '未修改成功']
                t = datetime.datetime.utcfromtimestamp(j.get('created_at')) + datetime.timedelta(hours=8)
                list1[0] = t.strftime("%Y-%m-%d %H:%M")
                list1[1] = str(j.get('id'))
                t = int(t.strftime("%Y%m%d%H%M"))
                if t > t2:
                    continue
                elif t < t1:
                    flag = 1
                    break
                list_n.append(list1)
            if flag == 1: break
        logger.info(f'获取到符合在"{setting["start_time"]}"到"{setting["end_time"]}"区间的有：{len(list_n)}条数据')
        return list_n

    def change_money(self, list_n):
        logger.info('开始修改金额')
        self.pdd_page.goto('https://yingxiao.pinduoduo.com/marketing/main/center/cpa/list')
        wait_page(self.pdd_page, seconds=10)
        time.sleep(15)
        n = 0
        if self.pdd_page.is_visible('//div[text()="跳过"]'):
            self.pdd_page.click('//div[text()="跳过"]')
        time.sleep(1)
        if self.pdd_page.is_visible('//span[@class="anticon drawer-close-icon"]'):
            self.pdd_page.click('//span[@class="anticon drawer-close-icon"]')
        for list1 in list_n:
            n = n + 1
            self.pdd_page.fill('//textarea[@id="id"]', list1[1])  # 填充要查询的id号 //span[@class="anticon drawer-close-icon"]
            self.pdd_page.click('//button[@class="anq-btn anq-btn-primary anq-btn-two-chinese-chars"]')  # 点击查询
            for t in range(7):
                text = self.pdd_page.text_content('//div[@class="GoodsTable_goodsRow__T_DIn"][1]')  # 获取id号
                if list1[1] in text:
                    break
                time.sleep(1)
            else: continue
            self.pdd_page.click('//span[@class="anticon BidPopover_icon__1M85N"][1]')  # 打开修改金额的界面
            self.pdd_page.click('//div[@class="BidPopover_input__3fHtQ"]//div[@class="IPT_inputBlockCell_17eihtg"]')  # 点击金额的输入框
            for dnum in range(10): self.pdd_page.keyboard.press('Backspace')  # 删去输入框的内容
            self.pdd_page.keyboard.insert_text(setting['money'])  # 输入要修改的金额
            list1[2] = setting['money']
            if setting['test_flag']:  # 点击取消
                self.pdd_page.click(
                    '//button[@class="anq-btn anq-btn-primary anq-btn-sm anq-btn-two-chinese-chars"]/../button[2]')
                list1[3] = '跑测成功'
                if n % 10 == 0: logger.info('正在跑测')
            else:  # 点击确定
                self.pdd_page.click('//button[@class="anq-btn anq-btn-primary anq-btn-sm anq-btn-two-chinese-chars"]')
                list1[3] = '修改成功'
                if n % 10 == 0: logger.info('正在修改')
        logger.info('修改完毕')
        return list_n

    def wait_for_selector(self, xpath, time1=20):
        try:
            self.pdd_page.wait_for_selector(xpath,timeout=1000*time1)
        except:
            time.sleep(1)

    def run(self):
        flag = self.login()
        if flag is False:
            return
        time.sleep(3)
        if self.pdd_page.is_visible('//div[@class="close-icon"]/i'):
            self.pdd_page.click('//div[@class="close-icon"]/i')  # 关闭弹窗//div[@class="close-icon"]/i
        self.wait_for_selector('//div[text()="在售中"]',time1=50)
        if self.pdd_page.is_visible('//span[@class="fail-title"]'):
            self.pdd_page.goto('https://yingxiao.pinduoduo.com/marketing/main/center/cpa/list')  # 刷新页面
        self.pdd_page.click('//div[text()="在售中"]')
        self.wait_for_selector('//tbody/tr[1]')
        list_n = self.get_data()
        if len(list_n) > 0:
            list_n = self.change_money(list_n)
            for i in list_n:
                self.data_dict['商品创建时间'].append(i[0])
                self.data_dict['商品ID'].append(i[1])
                self.data_dict['出价金额'].append(i[2])
                self.data_dict['修改结果'].append(i[3])
            df = pd.DataFrame(data=self.data_dict)
            df.to_excel(self.path, index=False)
            self.updata_file(self.path)
        else:
            logger.info(f'在{setting["start_time"]}到{setting["end_time"]}时间区间内的订单数量为零')


if __name__ == '__main__':
    # setting = {
    #     'acc':'乐言软件对接',
    #     'pwd':'Qitu123456',
    #     'start_time': '2022-02-10 10:25',
    #     'end_time': '2022-02-15 23:00',
    #     'money': '1.00',
    #     'test_flag': '11'
    # }
    if setting['start_time'] == '':
        setting['start_time'] = datetime.datetime.now().strftime('%Y-%m-%d 00:00')
    if setting['end_time'] == '':
        setting['end_time'] = datetime.datetime.now().strftime('%Y-%m-%d 23:59')
    if len(setting['start_time']) == 10:
        t1 = int(datetime.datetime.strptime(setting['start_time'], '%Y-%m-%d').strftime('%Y%m%d%H%M'))
        setting['start_time'] = datetime.datetime.strptime(setting['start_time'], '%Y-%m-%d').strftime('%Y-%m-%d %H:%M')
    elif len(setting['start_time']) == 16:
        t1 = int(datetime.datetime.strptime(setting['start_time'], '%Y-%m-%d %H:%M').strftime('%Y%m%d%H%M'))
    else:
        logger.info('日期格式错误，应为（2022-02-01）或（2022-02-01 10:01）')
    if len(setting['end_time']) == 10:
        t2 = int(datetime.datetime.strptime(setting['end_time'], '%Y-%m-%d').strftime('%Y%m%d2359'))
        setting['end_time'] = datetime.datetime.strptime(setting['end_time'], '%Y-%m-%d').strftime('%Y-%m-%d %H:%M')
    elif len(setting['end_time']) == 16:
        t2 = int(datetime.datetime.strptime(setting['end_time'], '%Y-%m-%d %H:%M').strftime('%Y%m%d%H%M'))
    else:
        logger.info('日期格式错误，应为（2022-02-01）或（2022-02-01 10:01）')
    now = int(datetime.datetime.now().strftime('%Y%m%d%H%M'))
    if t1 > t2:
        logger.info('开始时间应小于结束时间')
        raise 'error'
    elif t2 > now or t1 > now:
        logger.info('开始时间,结束时间应不大于今日')
        raise 'error'
    print(t1,t2)
    print(1)
    try:
        server = subprocess.Popen(
            r"./ms-playwright/chromium-907428/chrome-win/chrome.exe --remote-debugging-port=9002")
    except:
        os.popen(
            './chrome --no-sandbox --disable-gpu --disable-dev-shm-usage --use-gl=desktop --window-size=1600,1024 --remote-debugging-port=9002')
    time.sleep(2)
    p = sync_playwright().start()
    browser = p.chromium.connect_over_cdp('http://localhost:9002')
    context = browser.new_context(accept_downloads=False)
    page = context.new_page()
    try:
        pdd = PDD(page, setting['acc'], setting['pwd'])
        pdd.run()
    except Exception as e:
        logger.error(e)

    try:
        kill_chrome()
    except:
        server.kill()
