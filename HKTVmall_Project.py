from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import re
import pymysql


# 连接MYSQL数据库
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
        self.cursor.execute("DROP TABLE IF EXISTS HKTVMALL_TABLE")
        sql = """CREATE TABLE HKTVMALL_TABLE(
                Product_name VARCHAR(255),
                Sales_num INT(20),
                Original_price FLOAT(20),
                Activity_price FLOAT(20),
                Shop_name VARCHAR(255),
                Product_link NVARCHAR(1000) )"""
        self.cursor.execute(sql)
    # 存储提交数据
    def exec_data(self, sql, data=None):
        self.cursor.execute(sql, data)
        self.db.commit()
    # 魔术方法, 析构化 ,析构函数
    def __del__(self):
        self.cursor.close()
        self.db.close()

# 页面数据获取
def Get_data():
    # 注意区分find_elements_by_xpath和find_element_by_xpath
    Span_list = browser.find_elements_by_xpath("//*[@id='algolia-search-result-container']/div/span")
    for span in Span_list:
        # 商品名称
        Product_name = span.find_element_by_xpath('./div/div[2]/div[1]/div[1]/h4[1]').text
        # 销售数量，为空时默认为0
        try:
            Sales_num = span.find_element_by_xpath('./div/div[2]/div[1]/div[3]/span[1]').text
            Sales_num = int(re.search('([1-9]\d*\.?\d*)|(0\.\d*[1-9])',Sales_num.replace(',', '')).group())
        except:
            Sales_num = 0
        # 限时活动价
        try:
            Activity_price = span.find_element_by_xpath('./div/div[2]/div[2]/div[1]/div[2]/span[1]').text
            Activity_price = float(re.search('\$([1-9]\d*\.?\d*)|\$ ([1-9]\d*\.?\d*)',Activity_price).group().strip('$'))  #正则匹配后转化成float
        except:
            Activity_price = 0
        # 商品定价,存在不打折商品
        try:
            Original_price = span.find_element_by_xpath('./div/div[2]/div[2]/div[1]/div[1]/span[1]').text
            Original_price = float(re.search('\$([1-9]\d*\.?\d*)|\$ ([1-9]\d*\.?\d*)', Original_price).group().strip('$'))  #正则匹配后转化成float
        except:
            Original_price = Activity_price
        # 商店名称
        Shop_name = span.find_element_by_xpath('./div/div[2]/div[2]/a/span[1]').text
        # 商品链接
        Product_link = span.find_element_by_xpath('./div/a').get_attribute('href')
        # 存储数据
        insert_data = (
            "INSERT INTO HKTVMALL_TABLE(Product_name,Sales_num,Original_price,Activity_price,Shop_name,Product_link)" "VALUES(%s,%s,%s,%s,%s,%s)")
        product_data = (Product_name, Sales_num, Original_price, Activity_price, Shop_name, Product_link)
        MC.exec_data(insert_data, product_data)
    time.sleep(3)

# 页面翻滚
def Win_scroll():
    browser.execute_script('window.scrollTo(0,document.body.scrollHeight)')     # 翻滚到底
    time.sleep(2)

# 点击翻页
def Btn_click():
    btn = browser.find_element_by_xpath('//*[@id="paginationMenu_nextBtn"]')
    btn.click()
    time.sleep(4)


if __name__ == "__main__":
    chrome_options = Options()  # 创建配置对象
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    path = r'chromedriver.exe'
    browser = webdriver.Chrome(executable_path=path, chrome_options=chrome_options)
    browser.maximize_window()
    browser.get('https://www.hktvmall.com/hktv/zh/search_a/?keyword=%E7%B6%AD%E9%81%94&category=supermarket')
    time.sleep(1)

    # global All_yeshu  # 全局调用，爬取页数
    # browser.execute_script('window.scrollTo(0,document.body.scrollHeight)')
    All_yeshu = browser.find_element_by_xpath('//*[@id="search-result-wrapper"]/div/div[3]/div[3]/div/span[1]').text
    # browser.execute_script('window.scrollTo(0,0)')
    All_yeshu = int(re.search('([1-9]\d*\.?\d*)|(0\.\d*[1-9])', All_yeshu).group())

    MC = Mysqlconnect('127.0.0.1', 'root', '123456', 'hktvmall_database')
    MC.Cereate_table()
    Get_data()
    for i in range(All_yeshu-1):
        Win_scroll()
        Btn_click()
        Get_data()
        time.sleep(2)
    browser.quit()
