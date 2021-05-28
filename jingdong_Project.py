import re
import time
import random
import cv2
import pymysql
import numpy as np
from selenium import webdriver
from urllib import request
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains


# 链接MYSQL数据库
class Mysqlconnect(object):
    def __init__(self, host, user, password, database):
        '''
            :param host: IP
            :param user: 用户名
            :param password: 密码
            :param port: 端口号
            :param database: 数据库名
            :param charset: 编码格式
        '''
        self.db = pymysql.connect(host=host, user=user, password=password, port=3306, database=database, charset='utf8mb4')
        self.cursor = self.db.cursor()
    # 创建新表
    def Cereate_table(self):
        self.cursor.execute("DROP TABLE IF EXISTS JINGDONG_TABLE")
        sql = """CREATE TABLE JINGDONG_TABLE(
                    order_data VARCHAR(50),
                    order_num VARCHAR(30),
                    order_mode VARCHAR(30),
                    order_product VARCHAR(255),
                    product_num INT(10),
                    order_money FLOAT(20),
                    order_state VARCHAR(50) )"""
        self.cursor.execute(sql)
    # 存储提交数据
    def exec_data(self, sql, data=None):
        self.cursor.execute(sql, data)
        self.db.commit()
    # 关闭数据库
    def __del__(self):
        self.cursor.close()
        self.db.close()

# 页面翻滚到底部
def Win_scroll():
    browser.execute_script('window.scrollTo(0,document.body.scrollHeight)')
    time.sleep(2)

# 点击翻页
def Btn_click():
    btn = browser.find_element_by_xpath('//*[@id="paginationMenu_nextBtn"]')
    btn.click()
    time.sleep(3)

# 获取验证码与滑块之间相对位移
def Get_picture():
    # 获取链接
    back_url = browser.find_element_by_xpath('//*[@id="JDJRV-wrap-loginsubmit"]/div/div/div/div[1]/div[2]/div[1]/img').get_attribute('src')
    slider_url = browser.find_element_by_xpath('//*[@id="JDJRV-wrap-loginsubmit"]/div/div/div/div[1]/div[2]/div[2]/img').get_attribute('src')
    # 获取图片
    request.urlretrieve(back_url, 'back_pic')
    request.urlretrieve(slider_url, 'slider_pic')
    # 灰化图片
    back_pic2 = cv2.imread('back_pic', 0)
    slider_pic2 = cv2.imread('slider_pic', 0)
    # 保存灰化图片
    backName = 'back_pic2.jpg'
    sliderName = 'slider_pic2.jpg'
    cv2.imwrite(backName, back_pic2)
    cv2.imwrite(sliderName, slider_pic2)
    # 对滑块再度灰化
    block = cv2.imread(sliderName)
    block = cv2.cvtColor(block, cv2.COLOR_BGR2GRAY)
    block = abs(255 - block)
    cv2.imwrite(sliderName, block)
    # 将进终灰化过的图片读出来
    back = cv2.imread(backName)
    slider = cv2.imread(sliderName)
    # 获取偏移量
    result = cv2.matchTemplate(back, slider, cv2.TM_CCOEFF_NORMED)
    x, y = np.unravel_index(result.argmax(), result.shape)
    return int(y/1.3)   # 为什么除以1.3？？？

# 模拟滑动，速度变化
def btn_move(btn,x,y):
    All_distance = y
    Rel_distance = y
    Gone_distance = 0
    # 获取滑块
    element = browser.find_element_by_xpath(btn)                        # 定位滑块
    ActionChains(browser).click_and_hold(on_element=element).perform()  # 摁住滑块不动
    # 滑动过程考虑到更像人滑动，先快后慢，最后略微波动
    if All_distance < 80:
        while Rel_distance > 0:
            ratio = Rel_distance / All_distance
            if ratio < 0.2:
                # 结束阶段移动较慢
                span = random.randint(2, 5)
            elif 0.2 < ratio < 0.4:
                # 中间部分慢慢减速
                span = random.randint(7, 15)
            else:
                # 开始部分移动较快
                span = random.randint(20, 30)
            # 由于京东验证机制比较严格，模仿手动移动，每次移动上下有5像素的偏差
            ActionChains(browser).move_by_offset(span, random.randint(-5, 5)).perform()
            Rel_distance -= span
            Gone_distance += span
            # time.sleep(random.randint(5, 30) / 100)
        # 略微回移波动
        ActionChains(browser).move_by_offset(Rel_distance, random.randint(-3, 3)).perform()
        ActionChains(browser).release(on_element=element).perform()
    else:
        while Rel_distance > 0:
            ratio = Rel_distance / All_distance
            if ratio < 0.2:
                # 结束阶段移动较慢
                span = random.randint(4, 5)
            elif 0.2 < ratio < 0.4:
                span = random.randint(10, 20)
            else:
                # 开始部分移动较快
                span = random.randint(30, 40)
            # 由于京东验证机制比较严格，模仿手动移动，每次移动上下有5像素的偏差
            ActionChains(browser).move_by_offset(span, random.randint(-5, 5)).perform()
            Rel_distance -= span
            Gone_distance += span
            # time.sleep(random.randint(5, 30) / 100)
        # 略微回移波动
        ActionChains(browser).move_by_offset(Rel_distance, random.randint(-3, 3)).perform()
        ActionChains(browser).release(on_element=element).perform()
    time.sleep(2)


def Login():
    browser.find_element_by_xpath('//*[@id="ttbar-login"]/a[1]').click()
    time.sleep(2)
    browser.find_element_by_xpath('//*[@id="content"]/div[2]/div[1]/div/div[3]/a').click()
    time.sleep(1.5)
    inp_name = browser.find_element_by_xpath('//*[@id="loginname"]')
    inp_name.send_keys('有子有子')
    time.sleep(1.6)
    inp_password = browser.find_element_by_xpath('//*[@id="nloginpwd"]')
    inp_password.send_keys('sheishinibaba123')
    time.sleep(1.7)
    browser.find_element_by_xpath('//*[@id="loginsubmit"]').click()  # 京东账号密码登录必须滑动验证
    # 尝试多次登录验证
    for i in range(5):
        btn_distance = Get_picture()
        time.sleep(1.2)
        btn_move(r'//*[@id="JDJRV-wrap-loginsubmit"]/div/div/div/div[2]/div[3]', btn_distance, btn_distance)
        try:
            browser.find_element_by_xpath('//*[@id="loginsubmit"]').click()  # 如果捕获不到登录按键，说明已经成功跳转
        except:
            break
    print('登录成功')
    time.sleep(2)


def Get_order():
    All_order = browser.find_elements_by_xpath('/html/body/div[4]/div/div[1]/div[2]/div[4]/div[2]/table/tbody')
    for order in All_order:
        # 订单时间
        order_data = order.find_element_by_xpath('./tr[2]/td/span[2]').get_attribute('title')
        # 订单号
        order_num = order.find_element_by_xpath('./tr[2]/td/span[3]/a[1]').text
        # 支付方式
        order_mode = order.find_element_by_xpath('./tr[3]/td[3]/div/span[2]').text
        # 产品明细
        order_product = order.find_element_by_xpath('./tr[3]/td[1]/div[1]/div[2]/div[1]/a[1]').text
        # 产品数量
        product_num_text = order.find_element_by_xpath('./tr[3]/td[1]/div[2]').text
        product_num = int(re.search('([1-9]\d*\.?\d*)|(0\.\d*[1-9])', product_num_text).group())
        # 总金额
        order_money_text = order.find_element_by_xpath('./tr[3]/td[3]/div/span[1]').text
        order_money = float(re.search('\¥([1-9]\d*\.?\d*)|\¥ ([1-9]\d*\.?\d*)', order_money_text).group().strip('¥'))
        # 订单状态
        order_state_text = order.find_element_by_xpath('./tr[3]/td[4]/div/span[1]').text
        order_state = order_state_text.strip('')
        insert_data = (
            "INSERT INTO JINGDONG_TABLE(order_data, order_num, order_mode, order_product, product_num, order_money, order_state)" "VALUES(%s,%s,%s,%s,%s,%s,%s)")
        product_data = (order_data, order_num, order_mode, order_product, product_num, order_money, order_state)
        MC.exec_data(insert_data, product_data)


if __name__ == "__main__":
    chrome_options = Options()      # 创建配置对象
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])     # 以开发者模式
    path = r'C:\Users\zx303\PycharmProjects\new_chrome\chromedriver.exe'
    browser = webdriver.Chrome(chrome_options=chrome_options, executable_path=path)
    wait = WebDriverWait(browser, 3)
    browser.maximize_window()
    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument",
                           {"source": """ Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) """})
    browser.get('https://www.jd.com')
    time.sleep(2)
    Login()
    MC = Mysqlconnect('127.0.0.1', 'root', '123456', 'hktvmall_database')
    MC.Cereate_table()
    # 跳转订单页面,并切换
    browser.find_element_by_xpath('//*[@id="shortcut"]/div/ul[2]/li[3]/div/a').click()
    All_headle = browser.window_handles
    browser.switch_to.window(All_headle[1])
    time.sleep(2)
    Get_order()

    # 淘宝需扫码登录，无cookie
    # login_btn = browser.find_element_by_xpath('//*[@id="ttbar-login"]/a[1]').click()
    # time.sleep(10)
    # inp_search = browser.find_element_by_xpath('//*[@id="key"]')
    # inp_search.send_keys('枕头')
    # time.sleep(2)
    # btn_search = browser.find_element_by_xpath('//*[@id="search"]/div/div[2]/button/i').click()

    # 查找维达产品
    # inp_product = browser.find_element_by_id('//*[@id="q"]')
    # inp_product.send_keys('维达')
    # time.sleep(1)
    # browser.find_element_by_id('//*[@id="J_TSearchForm"]/div[1]/button').click()

    # Get_data()
    # for i in range(9):
    #     Win_scroll()
    #     Btn_click()
    #     Get_data()

    time.sleep(10)
    browser.quit()