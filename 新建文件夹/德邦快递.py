"""
德邦快递物流查询
"""


from playwright.sync_api import sync_playwright
from loguru import logger
from PIL import ImageChops
import os
import requests
import PIL.Image as Image
import numpy
import time
import json


executable_path1 = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
executable_path1 = executable_path1 if os.path.exists(executable_path1) else None


class DeBang:
    def __init__(self, id_str):
        self.order_id = id_str
        self.result_dict = {"data":[]}

    def append_table(self):
        """
        二值化过滤
        :return:
        """
        self.table = []
        for i in range(256):
            if i < 50:
                self.table.append(0)
            else:
                self.table.append(1)

    def login(self, playwright):
        self.browser = playwright.chromium.launch(headless=False,executable_path=executable_path1)
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        self.page.goto('https://www.deppon.com/index')

    def get_distance(self):
        """
        获取移动具体
        :return:
        """
        img_obj = self.page.wait_for_selector('//img[@id="grap_vacant"]')  # 查询到的页面等待加载
        style = img_obj.get_attribute('style')  # 访问style属性，获取字符串
        style_list = style.split(';')
        left = [i for i in style_list if 'left' in i][0]  # 查询特定的字符串
        distance = left.split(':')[-1].replace('px', '').replace(' ', '')
        return int(float(distance)) + 1

    def ease_out_quart(self, x):
        return 1 - pow(1 - x, 4)

    def get_tracks_2(self, distance, seconds, ease_func):
        """

        :param distance:
        :param seconds:
        :param ease_func:
        :return:
        """
        tracks = [0]
        offsets = [0]
        for t in numpy.arange(0.0, seconds, 0.1):
            ease = ease_func
            offset = round(ease(t / seconds) * distance)
            tracks.append(offset - offsets[-1])
            offsets.append(offset)
        tracks.extend([-3, -2, -3, -2, -2, -2, -2, -1, -0, -1, -1, -1])
        return tracks

    def move_to_gap(self, ix, iy, track ):
        """
        慢速移动鼠标
        :param ix:
        :param iy:
        :param track:
        :return:
        """
        mx = ix
        while track:
            x = track.pop(0)
            mx = mx + x
            self.page.mouse.move(mx, iy)
            time.sleep(0.05)
        self.page.mouse.up()

    def get_response(self):
        for i in range(5):
            try:
                distance = self.get_distance()  # 移动距离
                track = self.get_tracks_2(distance, 1, self.ease_out_quart)  # 移动距离2
                box = self.page.query_selector('//div[@id="swipper-btn"]/i').bounding_box()  # 图片边框微调
                with self.page.expect_response(f'https://www.deppon.com/gwapi/trackService/eco/track/searchTrack?billNo={self.order_id}') as response_info:
                    ix = box['x']
                    iy = box['y']
                    self.page.mouse.move(ix, iy)
                    self.page.mouse.down(button='left')
                    self.move_to_gap(ix, iy, track)
                response = response_info.value
                resp_json = response.json()
                return resp_json
            except Exception:
                time.sleep(1)
                self.page.click('//div[@class="actions"]/i[@class="iconfont iconRefresh"]')
                continue
        return False

    def get_type(self, resp_json):
        """
        获取状态
        :param resp_json:
        :return:
        """
        result = resp_json.get("result")
        type = result['billNoState']
        return type

    def process(self):
        """
        处理流程
        :return:
        """
        self.page.click('//textarea[@placeholder="请输入运单号查询"]')
        self.page.fill('//textarea[@placeholder="请输入运单号查询"]', self.order_id)
        self.page.press('//textarea[@placeholder="请输入运单号查询"]', 'Enter')  # 进入查询到的页面
        resp_json = self.get_response()
        if resp_json is False:
            self.result_dict['code'] = 1
            self.result_dict['msg'] = "查询失败"
            self.result_dict['data'] = "查询失败"
            return self.result_dict
        type1 = self.get_type(resp_json)
        data = {"order": self.order_id,
                "bill_state": type1}
        self.result_dict['code'] = 0
        self.result_dict['msg'] = "查询成功"
        self.result_dict['data'].append(data)

    def main(self, playwright):
        self.append_table()
        self.login(playwright)
        self.process()
        result_json = json.dumps(self.result_dict, ensure_ascii=False)
        return result_json


if __name__ == '__main__':
    id_str= "DPK364016052192"
    with sync_playwright() as playwright:
        db =DeBang(id_str)
        data_dict = db.main(playwright)
        print(data_dict)