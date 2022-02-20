"""
发货异常机器人
"""


import uuid
from playwright.sync_api import sync_playwright
import datetime
import time
import os
import json
import re
import pandas as pd
from loguru import logger
import utility

#
# setting ={
#     "login_url":"https://fxg.jinritemai.com/ffa/mshop/homepage/index",
#     "delivery_time":"2",  # 发货超时
#     "collection_time":"2",  #揽收超时
#     "transfer_time":"2",  # 中转超时
#     "dispatch_time":"2",  # 派签超时
#     "start_time": "2021-07-07",
#     "end_time" : "2021-07-08"
# }


class DouYinAbnormalDelivery:
    """
    发货异常
    """

    def __init__(self):
        self.login_url = setting['login_url']
        self.email = setting.get('email')
        self.password = setting.get('password')
        self.phone = setting.get('phone')
        self.data = []
        self.err_data = []
        self.order_id_list = []
        self.columns = ['买家昵称', '订单号', '付款时间', '发货时间', '物流公司', '物流单号', '物流最新节点', '发货时长（时）', '异常类型']
        self.excel_file = os.getcwd() + "\\data\\" + "{}发货异常表格.xlsx".format(str(uuid.uuid4()))
        self.err_file = os.getcwd() + "\\data\\" + "{}订单错误记录.xlsx".format(str(uuid.uuid4()))

    def login(self, playwright, is_retry_login=None):
        if is_retry_login is None:
            self.playwright = playwright
            import os
            if not os.path.exists("../UserData_firefox"):
                os.mkdir("../UserData_firefox", 755)
            self.context = self.playwright.firefox.launch_persistent_context(headless=False,
                                                                             user_data_dir='../UserData_firefox',
                                                                             executable_path='../ms-playwright/firefox-1234/firefox/firefox.exe',
                                                                             )

            self.login_page = self.context.pages[0]
        # 进入网页
        logger.info(' - 正在进入网页...')
        self.login_page.goto(self.login_url)
        logger.info(' - 开始登陆')
        self.wait(self.login_page)  # 自定义函数，等待网页加载
        logger.info(' - 网页已加载。')
        log_flg = ''
        for t in range(300, 0, -1):
            time.sleep(1)
            try:  # 尝试多种方法登录
                if self.login_page.is_visible('//*[@id="fxg-pc-header"]'):  # 找到登录成功才有的页面元素
                    logger.info(' - 账号已登陆。')
                    return True
                if t == 300:
                    if self.phone:
                        if self.email:
                            self.login_page.click('text="邮箱登录"')
                            self.login_page.click('//input[@placeholder="邮箱"]')
                            self.login_page.keyboard.insert_text(self.email)  # 模拟键盘，输入邮箱账号
                            self.login_page.click('//input[@placeholder="密码"]')
                            self.login_page.keyboard.insert_text(self.password)  # 模拟键盘，输入密码
                            self.login_page.click('//div[@class="account-center-submit"]/button')
                        else:
                            self.login_page.fill('//input[@placeholder="手机号码"]', self.phone)
                            # self.login_page.click("//div[@class='account-center-code-captcha']")
                    elif self.email:
                        self.login_page.click('text="邮箱登录"')
                        self.login_page.click('//input[@placeholder="邮箱"]')
                        self.login_page.keyboard.insert_text(self.email)
                        self.login_page.click('//input[@placeholder="密码"]')
                        self.login_page.keyboard.insert_text(self.password)
                        self.login_page.click('//div[@class="account-center-submit"]/button')
                    else:
                        self.login_page.click('//span[text()="抖音登录"]')
                        log_flg = '扫码'
                if t % 30 == 0:  # 扫码登录
                    if log_flg == '扫码':
                        if self.login_page.is_visible('//div[@class="btn-refresh"]'):
                            self.login_page.click('//div[@class="btn-refresh"]')
                            self.wait(self.login_page)  # 自定义函数，等待网页加载
                        logger.info(self.login_page.content())
                        img_base64 = self.login_page.get_attribute('//img[@class="qr"]', 'src')  # 获取二维码元素
                        img = f"""
                            <img class="qr" src="{img_base64}" alt="" style="width: 180px;height: 180px;">
                            """
                        logger.info(' - 请打开抖音APP,使用绑定抖音号扫码登录')
                        logger.info(img)
                    else:
                        pass
                if t % 5 == 0:
                    logger.info(' - 还未登录，请先登录，剩余%s秒' % t)
            except:  # 登录失败，下一轮循环执行
                continue
        else:
            logger.info(' - 登录超时，请重试！')
            raise False

    def wait(self, page, seconds=1):
        page.wait_for_load_state('networkidle')
        time.sleep(seconds)

    def remove_excess(self, page):
            self.wait(page)
            # 去掉通知
            if page.is_visible('//span[@class="ant-modal-close-x"]'):
                page.click('//span[@class="ant-modal-close-x"]')
            self.wait(page)
            # 去掉广告
            if page.is_visible('//div[@class="ant-modal-body"]'):
                page.click('//span[@aria-label="close-circle"]')
            self.wait(page)
            # 去掉引导
            if page.is_visible('text="新功能引导"'):
                page.click('//button[contains(.,"知道了")]')
                self.wait(page)
                page.click('//button[contains(.,"知道了")]')

    def choose_date(self):
        """
        选择日期
        :return:
        """
        self.wait(self.page)
        start_time = setting["start_time"] + " 00:00:00"
        end_time = setting["end_time"] + " 00:00:00"
        self.page.query_selector('//input[@id="compact_time"]').click()
        self.page.query_selector('//input[@id="compact_time"]').fill(start_time)
        self.page.click('//input[@placeholder="结束日期"]')
        self.page.fill('//input[@placeholder="结束日期"]', end_time)
        self.page.press('//input[@placeholder="结束日期"]', 'Enter')

    def get_order_id(self, data_list):
        """
        获取订单号
        :param date_list:
        :return:
        """
        for data in data_list:
            shop_order_id = data["shop_order_id"]
            self.order_id_list.append(shop_order_id)

    def get_post_addr(self, addr_dict):
        """
        获取收件地址
        :param addr_dict:
        :return:
        """
        province = addr_dict["province"]["name"]  # 省份
        city = addr_dict["city"]["name"]  # 市名
        town = addr_dict["town"]["name"]  # 区名
        street = addr_dict["street"]["name"]  # 街道
        detail = addr_dict["detail"]  # 街道
        addr = province + " " + city + ' ' + town + ' ' + street + ' ' +detail
        return addr

    def get_all_date(self):
        """
        获取时间
        :return:
        """
        try:
            date_list = []
            with self.page.expect_popup() as popup_info:
                self.page.click('//a[contains(text(),"订单详情")]')
            page1 = popup_info.value
            self.wait(page1)
            page1.wait_for_selector('//div[contains(text(),"买家下单")]')
            date_list_obj = page1.query_selector_all('//div[contains(@class,"style_description__")]')
            for date in date_list_obj:
                text = date.text_content()
                date_list.append(text)
        except Exception:
            pass
        finally:
            page1.close()
            return date_list

    def get_err_type(self, date1, hour):
        """
        获取异常类型
        :param date1: 付款时间/物流时间/其他时间
        :param hour: 给定时限
        :return: 异常类型
        """
        now = datetime.datetime.now()
        hour_int = int(hour)
        date1 = date1.replace('/','-')
        date1_date = datetime.datetime.strptime(date1, '%Y-%m-%d %H:%M:%S')
        date_difference = (now - date1_date).total_seconds()  # 时间差
        if date_difference > (hour_int * 3600):
            return True
        return False

    def append_data(self, data_info, delivery_time=None, company=None, shipment_id=None, latest_logistics_node=None, delivery_hour=None,
                    err_type=None):
        """
        添加数据
        :param data_info: 数据字典
        :param delivery_time: 发货时间
        :param company: 物流公司
        :param shipment_id: 物流单号
        :param latest_logistics_node: 最新节点
        :param delivery_hour: 发货时长
        :param err_type: 异常类型
        :return:
        """
        data_items = [
            data_info["nick_name"], data_info["id"], data_info["pay_time"], delivery_time, company, shipment_id, latest_logistics_node,
            delivery_hour, err_type]
        self.data.append(data_items)

    def beihuo_get_data_info(self, data, id):
        """
        获取详细数据
        :param data:
        :return:
        """
        order_info_dict = {}
        order_info_dict["id"] = id  # 订单号
        nick_name = data["user_nickname"] # 昵称
        order_info_dict["nick_name"] = nick_name
        # receiver_info = data["receiver_info"]
        # name = receiver_info["post_receiver"] # 收件人
        # order_info_dict["name"] = name
        # post_tel = receiver_info["post_tel"] # 收件人手机号
        # order_info_dict["post_tel"] = post_tel
        # post_addr = self.get_post_addr(receiver_info["post_addr"])  # 收件地址
        # order_info_dict["post_addr"] = post_addr
        date_list = self.get_all_date()  # 获取当前页面，特定元素的信息
        pay_time = date_list[1]  # 付款时间
        order_info_dict["pay_time"] = pay_time
        is_over = self.get_err_type(pay_time, setting["delivery_time"])  # 判断时间是否超时
        if is_over:
            order_info_dict["err_type"] = "发货超时"
            self.append_data(order_info_dict,err_type=order_info_dict["err_type"])  # 按格式输入内容到data中

    def beihuo_get_order_info(self):
        """
        获取订单信息
        :return:
        """
        for index, id in enumerate(self.order_id_list):
            try:
                logger.info(f'本次共处理{len(self.order_id_list)}条备货中的订单, 当前处理第{ index + 1 }条')
                self.page.fill('//input[@placeholder="请输入订单编号"]', str(id))
                with self.page.expect_response(f'**api/order/searchlist?order_id={id}**') as response_info:
                    self.page.click('text=\"查询\"')
                    response = response_info.value.text()
                    resp_json = json.loads(response)
                    data_list = resp_json.get("data")
                if len(data_list) >1:
                    logger.info("订单号{}匹配出多个订单详情".format(id))
                    err_msg = "订单号匹配出多个订单详情"
                    err_items = [str(id), err_msg]
                    self.err_data.append(err_items)
                    continue
                data_info = data_list[0]
                self.beihuo_get_data_info(data_info, id)  # 获取并填充data数据
            except Exception as e:
                logger.debug(e)
                continue

    def get_buyer_info(self, id):
        buyer_info_dict = dict()
        with self.page.expect_response(f'**/api/order/receiveinfo?come_from=pc&order_id={id}**') as response_info:
            self.page.click('//span[@data-kora="view"]')
            response = response_info.value.text()
            resp_json = json.loads(response)
            buyer_info = resp_json.get("data")
            nick_name = buyer_info["nick_name"]  # 昵称
            buyer_info_dict["nick_name"] = nick_name
            receive_info = buyer_info["receive_info"]
            post_receiver = receive_info["post_receiver"]  # 收件人
            buyer_info_dict["name"] = post_receiver
            post_tel = receive_info["post_tel"]  # 电话
            buyer_info_dict["post_tel"] = post_tel
            post_addr = self.get_post_addr(receive_info["post_addr"])  # 收件地址
            buyer_info_dict["post_addr"] = post_addr
            return buyer_info_dict

    def get_time(self, par_time_int):
        """
        数字转时间
        :param par_time_int:
        :return:
        """
        tup_time = time.localtime(par_time_int)
        sta_time = time.strftime("%Y/%m/%d %H:%M:%S", tup_time)
        return sta_time

    def get_latest_logistics_node(self, trace_list):
        """
        获取最新物流动态
        :param trace_list:
        :return:
        """
        trace_dict = dict()
        if len(trace_list) == 0:
            return False
        new_trace = trace_list[0]
        state_desc = new_trace.get("state_desc")  # 物流状态
        trace_dict["state_desc"] = state_desc
        desc = new_trace.get("desc")  # 物流节点信息
        trace_dict["desc"] = desc
        time = self.get_time(new_trace.get("time"))  # 状态时间
        trace_dict["time"] = time
        return trace_dict

    def get_delivery_hour(self, time1, time2):
        """
        获取发货时长
        :param time1: 发货时间
        :param time2: 付款时间
        :return:
        """
        delivery_time = time1.replace("/", "-")
        pay_time = time2.replace("/", "-")
        pay_time_datetime = datetime.datetime.strptime(pay_time, "%Y-%m-%d %H:%M:%S")
        delivery_time_datetime = datetime.datetime.strptime(delivery_time, "%Y-%m-%d %H:%M:%S")
        duration = (delivery_time_datetime - pay_time_datetime).total_seconds() / 3600
        return round(duration, 2)

    def get_err_code(self, trace_dict):
        state = trace_dict["state_desc"]
        if state in ["已揽件", "运输中"]:
            is_over = self.get_err_type(trace_dict["time"], setting["transfer_time"])
            if is_over:
                return "中转超时"
            return None
        if state == "派件中":
            is_over = self.get_err_type(trace_dict["time"], setting["dispatch_time"])
            if is_over:
                return "派件超时"
            return None

    def get_logistics_info(self, id):
        """
        获取物流信息
        :param id:
        :return:
        """
        logistics_info_dict = dict()
        with self.page.expect_response(f'**/api/order/getOrderLogistics?order_id={id}**') as response_info:
            self.page.click('//a[contains(text(),"查看物流")]')
            response = response_info.value.text()
            resp_json = json.loads(response)
            logistics_list = resp_json.get("data").get("logistics")
            logistics_info = logistics_list[0]
            logistics_name = logistics_info.get("logistics_name")  # 物流公司
            logistics_info_dict["logistics_name"] = logistics_name
            logistics_code = logistics_info.get("logistics_code")  # 运单号
            logistics_info_dict["logistics_code"] = logistics_code
            trace_list= logistics_info.get("trace")  # 物流动态列表
            trace_dict = self.get_latest_logistics_node(trace_list)  # 获取最新动态
            if trace_dict is False:  # 动态列表判断为零，输出False
                logistics_info_dict["latest_logistics_node"] = ""
            else:
                logistics_info_dict["latest_logistics_node"] = trace_dict["desc"]
            self.page.click('//button[@class="ant-btn ant-btn-primary"]/span[text()="确定"]')
            date_list = self.get_all_date()  # 获取所有的时间信息
            pay_time = date_list[1]  # 付款时间
            logistics_info_dict["pay_time"] = pay_time
            delivery_time = date_list[2]  # 发货时间
            logistics_info_dict["delivery_time"] = delivery_time
            delivery_hour = self.get_delivery_hour(delivery_time, pay_time)  # 发货时长
            logistics_info_dict["delivery_hour"] = delivery_hour
            if trace_dict is False:
                is_over = self.get_err_type(pay_time,setting["collection_time"])
                if is_over:
                    err_type = '揽收超时'
                    logistics_info_dict["err_type"] = err_type
                return logistics_info_dict
            err_type = self.get_err_code(trace_dict)
            if err_type is not None:
                logistics_info_dict['err_type'] = err_type
            return logistics_info_dict

    def fahuo_get_order_info(self):
        """
        发货信息获取
        :return:
        """
        for index, id in enumerate(self.order_id_list):
            try:
                url1 = 'https://fxg.jinritemai.com/ffa/morder/order/list'
                if self.page.url != url1:
                    self.page.goto(url1)
                    self.page.wait_for_load_state(state="networkidle")
                self.page.fill('//input[@placeholder="请输入订单编号"]', str(id))
                with self.page.expect_response(f'**api/order/searchlist?order_id={id}**') as response_info:
                    self.page.click('text=\"查询\"')
                    response = response_info.value.text()
                    resp_json = json.loads(response)
                    data_list = resp_json.get("data")
                if len(data_list) > 1:
                    logger.info("订单号{}匹配出多个订单详情".format(id))
                    err_msg = "订单号匹配出多个订单详情"
                    err_items = [str(id), err_msg]
                    self.err_data.append(err_items)
                    continue
                if len(data_list) == 0:
                    logger.info("订单号{}没有找到记录".format(id))
                    err_msg = "订单号没有找到记录"
                    err_items = [str(id), err_msg]
                    self.err_data.append(err_items)
                    continue
                # 获取用户信息
                # buyer_info_dict = self.get_buyer_info(id)
                # 获取物流信息
                logger.info(f'正在查询{id}单号的发货状态, 当前是{index+1}条, 共{len(self.order_id_list)}条')
                logistics_info_dict = self.get_logistics_info(id)  # 返回一个物流信息列表
                # 字典合并
                order_info_dict = {**logistics_info_dict}
                order_info_dict["id"] = id
                if "error_type" in order_info_dict.keys():
                    self.append_data(order_info_dict, delivery_time=order_info_dict["delivery_time"], company=order_info_dict["logistics_name"], shipment_id=order_info_dict["logistics_code"],
                                     latest_logistics_node= order_info_dict["latest_logistics_node"], delivery_hour=order_info_dict["delivery_hour"], err_type=order_info_dict["err_type"])
            except Exception as e:
                logger.debug(id, e)
                continue

    def write_excel(self):
        df = pd.DataFrame(data=self.data, columns=self.columns)
        # df.to_excel('aaaa.xlsx', index=False)
        df.to_excel(self.excel_file, index=False)

    def beihuo(self):
        """
        处理流程
        :return:
        """

        self.remove_excess(self.page)  # 自定义函数，去掉广告、引导等
        self.page.query_selector('//span[@title="10 条/页"]').click()
        self.page.query_selector('//div[contains(text(),"50 条/页")]').click()
        self.wait(self.page)
        self.page.query_selector('//div[@data-kora="展开"]').click()
        self.page.query_selector('//div[@data-kora="备货中"]').click()
        logger.info("当前正在处理备货中的订单")
        time.sleep(1)
        self.choose_date()  # 选择日期
        with self.page.expect_response('**/api/order/searchlist**&page=0**') as response_info:
            self.page.click('text=\"查询\"')
            response = response_info.value.text()  # 获取的response变为文本格式
            resp_json = json.loads(response)  # 变为json格式
            data_list = resp_json.get("data")
            if len(data_list) == 0:  # 无处理项，退出
                logger.info('当前日期范围内无可处理的备货订单')
                return
            self.wait(self.page, seconds=2)  # 有处理项，等待加载
            self.page.wait_for_selector('//li[@title="上一页"]')
            all_num1 = self.page.query_selector('//li[@title="上一页"]/preceding-sibling::li[1]').text_content()
            count_page1 = int(re.findall('\d+', all_num1)[0]) // 50 + 1  # 对获取的文本数字进行处理,得到有多少页
            self.get_order_id(data_list)  # 把获取的订单号输入到order_id_list中
        for i in range(1, count_page1):
            with self.page.expect_response(f'**/api/order/searchlist**&page={i}**') as response_info:
                self.page.click('//li[@title="下一页"]')
                response = response_info.value.text()
                resp_json = json.loads(response)
                data_list = resp_json.get("data")
                self.get_order_id(data_list)
        self.beihuo_get_order_info()

    def fahuo(self):
        """
        发货
        :return:
        """
        self.page.query_selector('//div[@data-kora="已发货"]').click()  # 点击发货
        logger.info("开始处理已发货订单")
        self.wait(self.page)
        self.choose_date()  # 选择日期
        with self.page.expect_response('**/api/order/searchlist?order_status=on_delivery**page=0**')as response_info:
            self.page.click('text=\"查询\"')
            self.wait(self.page)
            all_num2 = self.page.query_selector('//li[@title="上一页"]/preceding-sibling::li[1]').text_content()
            count_page2 = int(re.findall('\d+', all_num2)[0]) // 50 + 1  # 获取有多少页
            response = response_info.value.text()
            resp_json = json.loads(response)
            data_list = resp_json.get("data")
            self.get_order_id(data_list)
        for i in range(1, count_page2):
            with self.page.expect_response(f'**/api/order/searchlist?order_status=on_delivery**page={i}**') as response_info:
                self.page.click('//li[@title="下一页"]')
                self.wait(self.page)
                response = response_info.value.text()
                resp_json = json.loads(response)
                data_list = resp_json.get("data")
                self.get_order_id(data_list)

    def main(self, playwright):
        self.login(playwright)
        self.page = self.context.new_page()
        self.page.goto('https://fxg.jinritemai.com/ffa/morder/order/list')
        self.beihuo()
        self.order_id_list = []
        self.fahuo()
        self.fahuo_get_order_info()
        self.write_excel()
        file_name = self.excel_file.split("\\")[-1]
        utility.post_file(file_name)  # 用POST方法上传文件
        out_url = utility.get_download_url(file_name)  # 定义文件的下载方法
        logger.info(f'任务完成, 请下载<a href="{out_url}" target="_blank">发货异常文件</a>，查看格式是否正确')


fh = DouYinAbnormalDelivery()
with sync_playwright() as playwright:
    logger.info("*"*10 + "任务开始" + "*"*10)
    fh.main(playwright)
    logger.info("*" * 10 + "任务完毕" + "*" * 10)