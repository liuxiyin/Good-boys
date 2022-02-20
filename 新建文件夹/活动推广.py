"""
抖店-活动推广机器人
"""

import time
from playwright.sync_api import sync_playwright
from loguru import logger
import pandas as pd
import utility
import uuid
import os
import re
import json
import datetime


setting ={
    "login_url":"https://fxg.jinritemai.com/ffa/mshop/homepage/index",
    "start_time": "2021-05-30",
    "end_time" : "2021-05-31"
}


class PromoteRobot:
    """
    推广机器人
    """

    def __init__(self):
        self.order_id_list = []
        self.content_list = []
        self.data = []
        self.columns = ["日期", "订单号", "营销话术", "发送结果"]
        self.excel_file = os.getcwd() + "\\data\\" + "{}.xlsx".format(str(uuid.uuid4()))

    def login(self, playwright):
        self.context = playwright.firefox.launch_persistent_context(headless=False, user_data_dir='UserData_firefox',
                                                                    # executable_path='../ms-playwright/firefox-1234/firefox/firefox.exe',
                                                                    )
        self.login_page = self.context.pages[0]
        self.login_page.goto(setting['login_url'])
        for t in range(300, 0, -1):
            time.sleep(1)
            if self.login_page.is_visible('//h1[contains(text(),"抖店")]'):
                logger.info(' - 账号已登陆。')
                return True

    def recoding(self, id, msg="发送成功"):
        """
        添加记录
        :param id: 订单号
        :param msg: 发送结果
        :return:
        """
        today = datetime.datetime.today().date()
        date_items = [str(today), str(id), setting["content"], msg]
        self.data.append(date_items)

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

    def query_content_id(self):
        """
        查询有评论的订单
        :return:
        """
        logger.info('正在查询订单是否有评价-------')
        for index, id in enumerate(self.order_id_list):
            self.page1.fill('//input[@placeholder="商品编号/订单编号/商品名称"]', str(id))
            with self.page1.expect_response(f'**/product/tcomment/commentList?**id={id}**') as response_info:
                self.page1.click('//span[contains(text(),"查询")]')
                response = response_info.value.text()
                resp_json = json.loads(response)
                data_list = resp_json.get("data")
                if len(data_list) == 0:
                    logger.info(f'{id}无评价内容')
                    logger.info(f"共{len(self.order_id_list)}条订单, 当前查询第{index + 1}条")
                    continue
                data_info = data_list[0]
                content = data_info["content"]
                self.content_list.append(id)

    def send_files(self):
        """
        发送信息
        :return:
        """
        logger.info('----正在发送推送信息----')
        for index, id in enumerate(self.content_list):
            logger.info(f'共{len(self.content_list)}条订单需要发送,当前发送第{ index + 1}条')
            self.page.fill('//input[@placeholder="搜用户/180天订单(Ctrl+F)"]', id)
            self.page2.click('//div[contains(@class,"TdhgxWD_")]')
            if self.page2.is_visible('//div[contains(text(),"会话超过7天不可回复")]'):
                err_msg = "会话超过7天不可回复"
                self.recoding(id, msg=err_msg)
                continue
            self.page2.fill('//textarea', setting["msg"])
            self.page2.click('//div[contains(text(),"发送")]')
            # 发送图片
            self.page2.set_input_files('//textarea/preceding-sibling::div[1]/div/label/input', setting["image_path"])
            div_list = self.page2.query_selector('//div[@id="root"]/following-sibling::div[4]/div/div[3]')
            for div in div_list:
                text = div.text_content()
                if text == "发送":
                    div.click()
            # 发送视频
            if "video_path" in setting.keys() or setting["video_path"] is not None:
                self.page2.set_input_files('//textarea/preceding-sibling::div[1]/div/label[2]', '')
                div_list = self.page2.query_selector('//div[@id="root"]/following-sibling::div[4]/div/div[3]')
                for div in div_list:
                    text = div.text_content()
                    if text == "发送":
                        div.click()

            """
            关闭会话
            """
            self.recoding(id)

    def process(self):
        """
        处理流程
        :return:
        """
        self.page.query_selector('//span[@title="10 条/页"]').click()
        self.page.query_selector('//div[contains(text(),"50 条/页")]').click()
        self.wait(self.page, seconds=2)
        self.page.query_selector('//div[@data-kora="展开"]').click()
        self.page.click('//span[@title="下单时间"]')
        self.page.click('//div[contains(text(),"完成时间")]')
        self.choose_date()  # 输入日期
        with self.page.expect_response('**/api/order/searchlist**&page=0**') as response_info:
            self.page.click('text=\"查询\"')
            self.wait(self.page)
            all_num1 = self.page.query_selector('//li[@title="上一页"]/preceding-sibling::li[1]').text_content()
            count_page1 = int(re.findall('\d+', all_num1)[0]) // 50 + 1
            print("count_page1", count_page1)
            response = response_info.value.text()
            resp_json = json.loads(response)
            data_list = resp_json.get("data")
            self.get_order_id(data_list)
        for i in range(1, count_page1):
            with self.page.expect_response(f'**/api/order/searchlist**&page={i}**') as response_info:
                self.page.click('//li[@title="下一页"]')
                response = response_info.value.text()
                resp_json = json.loads(response)
                data_list = resp_json.get("data")
                self.get_order_id(data_list)
                time.sleep(1)

    def write_excel(self):
        """
        生成表格
        :return:
        """
        df = pd.DataFrame(data=self.data, columns=self.columns)
        df.to_excel(self.excel_file, index=False)

    def main(self, playwright):
        self.login(playwright)
        self.page = self.context.new_page()
        self.page.goto('https://fxg.jinritemai.com/ffa/morder/order/list')  # 进入订单管理页面
        self.login_page.close()
        self.remove_excess(self.page)  # 去掉广告
        self.process()
        self.page1 = self.context.new_page()
        self.page1.goto('https://fxg.jinritemai.com/ffa/g/comment')  # 进入评价管理页面
        self.page.close()
        self.remove_excess(self.page1)
        self.query_content_id()
        self.page1.close()
        df = pd.DataFrame(data=self.content_list, columns=["订单号"])
        df.to_excel(self.excel_file, index=False)
        self.page2 = self.context.new_page()
        self.page2.goto('https://im.jinritemai.com/pc_seller/main/chat')  # 进入飞鸽系统
        self.remove_excess(self.page2)
        self.send_files()



if __name__ == '__main__':
    pr = PromoteRobot()
    with sync_playwright() as playwright:
        pr.main(playwright)
