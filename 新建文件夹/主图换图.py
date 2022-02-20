"""
乐言主图换图机器人
"""

from playwright.sync_api import sync_playwright
from loguru import logger
import os
import pandas as pd
import time


setting = {
    'url': "https://sucai.wangpu.taobao.com",
    "main_image_path": "全部图片/天猫活动/双11预热/主图/",  # 主图路径
    "leng_image_path": "全部图片/天猫活动/双11预热/长图",  # 长图路径
    "account": "loveebuyer",
    "pwd": "Zpf19870501.."
}



executable_path = "../ms-playwright/firefox-1281/firefox/firefox.exe"


class ChangeImage:
    def __init__(self):
        self.main_image_id_list = []
        self.leng_image_id_list = []

    def login(self, playwright):
        self.browser = playwright.chromium.connect_over_cdp("http://localhost:9002")
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        self.page.goto(setting.get('url'))
        self.page.click('//div[@id="login"]/div/i[@class="iconfont icon-qrcode"]')
        for i in range(300, 0, -1):
            try:
                if self.page.is_visible('//div[@id="wp-header"]'):
                    logger.info('登录成功')
                    return True
                if i % 30 == 0:
                    img_64 = self.page.query_selector('//div[@class="qrcode-img"]').screenshot(path="data/log.png")
                    img = f""" <img class="qr" src="{img_64}" alt="" style="width: 180px;height: 180px;"> """
                    logger.info(img)
                    logger.info('请打开淘宝APP进行扫码登录')
                if self.page.is_visible('//div[@class="qrcode-error"]'):
                    self.page.click('//button[@class="refresh"]')
            except Exception:
                continue
            finally:
                time.sleep(1)
        return False

    def remve_alert(self):
        """
        删除弹窗
        :return:
        """
        # self.page.wait_for_load_state(state="networkidle")
        if self.page.is_visible('//div[@class="guideSkip" and text()="跳过"]'):
            self.page.click('//div[@class="guideSkip" and text()="跳过"]')
        if self.page.is_visible('//a[text()="知道了"]'):
            self.page.click('//a[text()="知道了"]')

    def choose_image_path(self, image_list1, image_list2):
        """
        选择图片路径
        :param image_list:
        :return:
        """
        for i in range(len(image_list1)):
            if i != (len(image_list1)-1) and self.page.is_visible(f'//li/a[@title="{image_list1[i+1]}"]') is False:
                self.page.query_selector(f'//li/a[@title="{image_list1[i]}"]').dblclick()  # 双击图片元素
                time.sleep(0.1)
                continue
        self.page.click(f'//li/a[@title="{image_list1[-1]}"]')
        time.sleep(1)
        main_obj = self.page.query_selector_all('//div[@class="block-mid"]/div[@class="block-mid-lis "]/div[@class="mid-lis-name"]')
        self.get_image_id(main_obj, self.main_image_id_list)
        self.page.click(f'//li/a[@title="{image_list2[-1]}"]')
        time.sleep(1)
        leng_obj = self.page.query_selector_all('//div[@class="block-mid"]/div[@class="block-mid-lis "]/div[@class="mid-lis-name"]')
        self.get_image_id(leng_obj, self.leng_image_id_list)


    def get_image_id(self, obj_list, id_list):
        """
        获取id
        :return:
        """

        for i in obj_list:
            image_name = i.get_attribute('title')
            image_id = image_name.split(".png")[0]
            id_list.append(image_id)

    def update_image_main(self, id, page, main_path_list, ):
        """
        新页面替换图片
        :param page:
        :return:
        """
        main_image_obj = page.query_selector(
            '//label[text()="电脑端宝贝图片"]/parent::div[1]/following-sibling::div//div[@class="info-content"]//div/p[text()="宝贝主图"]')
        # 删除图片
        if main_image_obj.query_selector('//following-sibling::div[2]/div[@class="image-card"]'):
            main_image_obj.query_selector('//following-sibling::div[2]/div[@class="image-card"]/img').hover()
            main_image_obj.query_selector('//following-sibling::div[2]/div[@class="image-tools"]/i[contains(@class,"remove")]').click()
        # 上传图片
        if not page.query_selector('//iframe[contains(@src, "//sucai.wangpu.taobao.com/select.htm?")]'):
            main_image_obj.query_selector('//following-sibling::div[@class="image-upload-btn"]/div/p').click()
        image_frame = page.wait_for_selector('//iframe[contains(@src, "//sucai.wangpu.taobao.com/select.htm?")]').content_frame()
        time.sleep(2)
        for i in range(5):
            if image_frame.query_selector('//div[@class="next-select-inner"]/input').get_attribute('value') == main_path_list[-2]:
                break
            image_frame.fill('//div[@class="next-select-inner"]/input', main_path_list[-2])
            image_frame.press('//div[@class="next-select-inner"]/input', 'Enter')
            time.sleep(2)
        image_frame.query_selector(f'//li/a[@title="{main_path_list[-2]}"]').dblclick()
        image_frame.query_selector(f'//li/a[@title="{main_path_list[-1]}"]').click()
        time.sleep(1)
        image_list = image_frame.query_selector_all('//div[@id="items"]//div[@class="name"]')
        for item in image_list:
            name =item.get_attribute('title')
            if id in name:
                item.click()
                page.click('//label[text()="电脑端宝贝图片"]')
                page.click('//button[@id="button-submit"]')
                page.close()
                return True
        return False

    def update_image_leng(self, id, new_page, leng_path_list):
        """
        更换长图
        :param id:
        :param page:
        :param leng_path_list:
        :return:
        """
        len_image_obj = new_page.query_selector(
            '//label[text()="宝贝长图"]/parent::div[1]/following-sibling::div//div[@class="info-content"]//div/p[text()="宝贝长图"]')
        # 删除旧图
        if len_image_obj.query_selector('//following-sibling::div[2]/div[@class="image-card"]'):
            len_image_obj.query_selector('//following-sibling::div[2]/div[@class="image-card"]/img').hover()
            len_image_obj.query_selector('//following-sibling::div[2]/div[@class="image-tools"]/i[contains(@class,"remove")]').click()
        # 上传图片
        len_image_obj.query_selector('//following-sibling::div[@class="image-upload-btn"]/div/p').click()
        image_frame = new_page.query_selector('//iframe[contains(@src, "//sucai.wangpu.taobao.com/select.htm?")]').content_frame()
        image_frame.fill('//div[@class="next-select-inner"]/input', leng_path_list[-2])
        image_frame.press('//div[@class="next-select-inner"]/input', 'Enter')
        if image_frame.is_visible(f'//li/a[@title="{leng_path_list[-1]}"]') is False:
            image_frame.query_selector(f'//li/a[@title="{leng_path_list[-2]}"]').click()
        image_frame.click(f'//li/a[@title="{leng_path_list[-1]}"]')
        image_list = image_frame.query_selector_all('//div[@id="items"]//div[@class="name"]')
        for item in image_list:
            name =item.get_attribute('title')
            if id in name:
                item.click()
                new_page.click('//label[text()="电脑端宝贝图片"]')
                break
        if not image_frame.query_selector('//div[@class="jcrop-holder"]/div[@ng-mouseover="resizeShow = true"]/div/div[@class="jcrop-tracker"]'):
            return False
        # box_size = image_frame.query_selector('//div[@class="jcrop-holder"]/div[@ng-mouseover="resizeShow = true"]/div/div[@class="jcrop-tracker"]')
        # self.page.mouse.down(button="right")
        big_image = image_frame.query_selector('//div[@class="jcrop-holder"]')
        simall_image = image_frame.query_selector('//div[@class="jcrop-holder"]/div[@ng-mouseover="resizeShow = true"]/div/div[@class="jcrop-tracker"]')
        box1 = big_image.bounding_box()
        box2 = simall_image.bounding_box()
        new_page.mouse.move(box2["x"]+1, box2["y"]+1)
        new_page.mouse.down()
        new_page.mouse.move(box1["x"]+1, box1["y"]+1)
        new_page.mouse.up()
        image_frame.fill('//div[@class="crop-value"]/div[1]/span/input', '400')
        image_frame.fill('//div[@class="crop-value"]/div[2]/span/input', '600')
        image_frame.press('//div[@class="crop-value"]/div[2]/span/input', 'Enter')
        new_page.click('//button[@id="button-submit"]')

    def change_image(self, main_path_list, leng_path_list):
        """
        替换图片
        :return:
        """
        try:
            self.page.goto('https://item.manager.taobao.com/taobao/manager/render.htm', timeout=10000)
        except Exception:
            pass
        self.remve_alert()
        for i in self.main_image_id_list:
            self.page.fill('//input[@name="queryItemId"]',i)
            self.page.click('//button[text()="查询"]')
            with self.page.expect_popup() as popup_info:
                self.page.click('//button[@title="编辑商品"]')
            new_page = popup_info.value
            new_page.wait_for_load_state(state="networkidle")
            main_flg = self.update_image_main(i, new_page, main_path_list)
            if main_flg is False:
                logger.info(f"{i}在主图中查询不到")
            time.sleep(1)
            # leng_flg = self.update_image_leng(i, new_page, leng_path_list)


    def process(self):
        """
        业务处理流程
        :return:
        """
        self.remve_alert()
        image_path_main = setting.get('main_image_path')
        image_path_leng = setting.get('leng_image_path')
        main_path_list = image_path_main.split('/')  # 主图路径切割成多个字符串
        main_path_list = [i for i in main_path_list if i!=""]
        leng_path_list = image_path_leng.split('/')
        leng_path_list = [i for i in leng_path_list if i!=""]
        self.choose_image_path(main_path_list, leng_path_list)
        self.change_image(main_path_list, leng_path_list)


    def main(self, playwright):

        log_flg = self.login(playwright)
        if log_flg is False:
            logger.info('登录失败，请重新启动机器人')
        self.process()


if __name__ == '__main__':

    import subprocess
    # 打开浏览器
    service = subprocess.Popen("ms-playwright/chromium-907428/chrome-win/chrome --remote-debugging-port=9002")  # 执行cdp链接防反爬，可以看二维码
    time.sleep(2)
    with sync_playwright() as playwright:
        change_image = ChangeImage()
        change_image.main(playwright)
    service.kill()  # 关闭浏览器
