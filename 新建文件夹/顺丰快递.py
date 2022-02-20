from playwright.sync_api import sync_playwright
from loguru import logger
from PIL import ImageChops
import os
import requests
import PIL.Image as Image
import numpy
import time
import arrow

executable_path1 = "./ms-playwright/chromium-901522/chrome-win/chrome.exe"
executable_path1  =executable_path1  if os.path.exists(executable_path1) else None


class ShunFengOrder:
    """顺丰订单"""
    def __init__(self, id_str):
        self.order_id = id_str
        self.result_dict = {}
        self.type_dict = {"1":"未出库", "2":"已揽收", "3":"运输中", "4":"已签收"}

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
        self.browser = playwright.chromium.launch(headless=False,
                                                  executable_path=executable_path1
                                                  )
        self.context = self.browser.new_context()
        self.page = self.context.new_page()

    def ease_out_quart(self, x):
        return 1 - pow(1 - x, 4)

    def get_background(self):
        """
        获取背景图
        :return:
        """
        frame = self.page.wait_for_selector('//iframe[@id="tcaptcha_popup"]').content_frame()  # 标签页获取
        base_ulr = "https://captcha.guard.qcloud.com"
        bg = frame.wait_for_selector('//img[@id="slideBkg"]')  # 等待标签页的元素
        bg_src = bg.get_attribute('src')  # 图片的 Xpath 地址
        bg_url = base_ulr + bg_src  # 网页上的图片地址
        response = requests.get(bg_url)
        with open('data/11.png', 'wb') as f:
            f.write(response.content)
        bg_img = Image.open('data/11.png')
        bg_img = bg_img.resize((340,195))
        bg_all = base_ulr + (bg_src[:-1]+'0')
        response = requests.get(bg_all)
        with open('data/22.png', 'wb') as f:
            f.write(response.content)
        bg_all_img = Image.open('data/22.png')
        bg_all_img = bg_all_img.resize((340,195))
        return bg_img, bg_all_img

    def compute_gap(self, img1, img2):
        """计算缺口偏移 这种方式成功率很高"""
        # 将图片修改为RGB模式
        img1 = img1.convert("RGB")
        img2 = img2.convert("RGB")

        # 计算差值
        diff = ImageChops.difference(img1, img2)
        # 灰度图
        diff = diff.convert("L")
        # 二值化
        diff = diff.point(self.table, '1')
        for w in range(diff.size[0]):
            lis = []
            for h in range(diff.size[1]):
                if diff.load()[w, h] == 1:
                    lis.append(w)
                if len(lis) > 5:
                    return w

    def get_tracks_2(self, distance, seconds, ease_func):
        """
        获取已定距离矩阵
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

    def remove_verification(self):
        """
        顺丰滑块验证
        :return:
        """
        frame = self.page.wait_for_selector('//iframe[@id="tcaptcha_popup"]').content_frame()
        if frame.wait_for_selector('//img[@id="slideBkg"]') is None:
            self.browser.new_context()
        time.sleep(1)
        for i in range(5):
            try:
                with self.page.expect_response(f'https://www.sf-express.com/sf-service-core-web/service/waybillRoute/{self.order_id}**') as response_info:
                    bg_img, car_img = self.get_background()
                    dist = self.compute_gap(bg_img,car_img)  # 计算图片距离
                    track = self.get_tracks_2(dist, 1, self.ease_out_quart)  # 计算距离
                    box = frame.wait_for_selector('//div[@id="tcaptcha_drag_button"]')
                    box_size = box.bounding_box()  # 边框微调，获取坐标
                    ix = box_size['x'] + 34
                    iy = box_size['y'] + 19
                    self.page.mouse.move(ix, iy)  # 鼠标移动到滑块上
                    self.page.mouse.down(button='left')  # 按下鼠标左键
                    self.move_to_gap(ix, iy, track)  # 慢速移动鼠标
                    time.sleep(1)
                response = response_info.value
                resp_json = response.json()
                return resp_json
            except Exception:
                continue
        return False

    def get_order_type(self, resp_json):
        """
        获取物流时间
        :param date_time:
        :return:
        """
        result = resp_json.get("result")
        routes = result.get('routes')
        if len(routes) == 0:
            return False
        type = self.type_dict.get(routes[0]['billFlag'])
        return type

    def process(self):
        self.result_dict['id'] = self.order_id
        self.page.goto(f'https://www.sf-express.com/cn/sc/dynamic_function/waybill/#search/bill-number/{self.order_id}')
        self.page.wait_for_load_state(state="networkidle")
        try:
            resp_json = self.remove_verification()
            if resp_json is False:
                raise Exception("失败")
            order_type = self.get_order_type(resp_json)
            if order_type is False:
                raise Exception("运单号错误或重复")
            self.result_dict["code"] = 0
            self.result_dict['msg'] = "查询成功"
            self.result_dict['data'] = order_type
        except Exception as e:
            self.result_dict['code'] = 1
            self.result_dict['msg'] = '查询失败'
            self.result_dict['data'] = e

    def main(self, playwright):
        self.append_table()
        self.login(playwright)
        self.process()
        self.browser.close()
        return self.result_dict


if __name__ == '__main__':
    order_str = "SF1327660951871"
    with sync_playwright() as playwright:
        sf = ShunFengOrder(order_str)
        data_dict = sf.main(playwright)
        print(data_dict)
