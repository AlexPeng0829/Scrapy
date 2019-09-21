from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import xlwt



class web_site:

    #对象初始化
    def __init__(self):
        url = 'http://oa.n-huashan.org.cn:58888/OADemo/Index.aspx'
        self.url = url
        self.row_to_append = 0
        self.header_exist = False

        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2}) # 不加载图片,加快访问速度
        options.add_experimental_option('excludeSwitches', ['enable-automation']) # 此步骤很重要，设置为开发者模式，防止被各大网站识别出来使用了Selenium

        self.browser = webdriver.Chrome(executable_path=chromedriver_path, options=options)
        self.wait = WebDriverWait(self.browser, 10) #超时时长为10s


    def login(self):
        self.browser.get(self.url)


    def iterate_over_date(self):
        #TODO Now only iterate over September
        dates = self.date_generator()
        # dates = [['2018-09-16'], ['2019-09-16', '2019-09-17']]
        workbook = xlwt.Workbook()
        # sheet = workbook.add_sheet("sheet_name") 
        # for each_date in dates:
        #     self.search(each_date)
        #     self.parse_data(each_date, sheet)
        # workbook.save("D:\\code\\data.xls")

        for dates_per_year in dates:
            sheet = workbook.add_sheet(dates_per_year[0][0:4])
            self.header_exist = False
            for each_date in dates_per_year:
                self.search(each_date)
                self.parse_data(each_date, sheet)
        workbook.save("D:\\code\\data.xls")

    def date_generator(self):
        dates = list()
        for year in range(2018, 2020):
            dates_per_year = list()
            for month in range(1, 13):
                if month in (1, 3, 5, 7, 8, 10 ,12):
                    end_day = 32
                elif month ==2:
                    end_day = 29
                elif month in (4, 6, 9, 11):
                    end_day = 31
                else:
                    print("Value error")
                    # raise(ValueError)
                for day in range(1, end_day):
                    date = str(year) + '-' + str(month).rjust(2, '0') + '-' + str(day).rjust(2, '0')
                    dates_per_year.append(date)
            dates.append(dates_per_year)
        return dates
                
    def search(self, search_date):
        self.browser.find_element_by_id('txtRQ').clear()
        self.browser.find_element_by_id('txtRQ').send_keys(search_date)
        self.browser.find_element_by_name("ddlSSLC").send_keys("手术室、门诊手术室")
        self.browser.find_element_by_name("Btn_Search").click()

    
    def parse_data(self, search_date, sheet):
        soup = BeautifulSoup(self.browser.page_source,'lxml')
        tables = soup.find_all("table")
        tab_with_real_data = tables[1]

        if self.header_exist == False:
            for tr in tab_with_real_data.findAll('tr'):
                sheet.write(self.row_to_append, 0, search_date)
                col = 1
                for td in tr.findAll('td'):
                    # print (td.getText())
                    sheet.write(self.row_to_append, col, td.getText())
                    col+=1
                self.row_to_append+=1
            self.header_exist = True
        else:
            for tr in tab_with_real_data.findAll('tr')[1:]:
                sheet.write(self.row_to_append, 0, search_date)
                col = 1
                for td in tr.findAll('td'):
                    sheet.write(self.row_to_append, col, td.getText())
                    col+=1
                self.row_to_append+=1


    # def save_data(self):
        



if __name__ == "__main__":
    

    chromedriver_path = "C:\\Users\\jiasu\\Downloads\\chromedriver_win32\\chromedriver.exe" 

    a = web_site()
    a.login() 
    a.iterate_over_date()
