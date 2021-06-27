# Load packages
from selenium import webdriver
from time import sleep
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.keys import Keys
import os
import datetime

# disable visualization
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
# avoid detection
options = ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-automation'])
# set default download directory
wd = os.getcwd()
download_dir = os.path.join(wd,'output')
prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': download_dir}
chrome_options.add_experimental_option('prefs', prefs)

# instantiate browser object
browser = webdriver.Chrome(executable_path='./chromedriver.exe',
                           chrome_options=chrome_options,
                           options=options)



def log_in():
    print('The crawler is started...')
    try:
        log_in_button = browser.find_element_by_xpath('//*[@id="header___3l7sS"]/div[2]/span[1]')
        log_in_button.click()
        userName_tag = browser.find_element_by_id('_login_email')
        password_tag = browser.find_element_by_id('_login_password')
        remember_login = browser.find_element_by_id('_login_remember')
        sleep(1)
        userName_tag.send_keys('xinmei.yang@ruc.edu.cn')
        sleep(1)
        password_tag.send_keys('tvzone123456')
        sleep(1)
        remember_login.click()
        sleep(1)
        btn = browser.find_element_by_class_name('ant-btn-lg')
        btn.click()
        sleep(3)
        print('Login Successfully...')
    except:
        print('Login fail, please try again...')

def getsource():
    try:
        JieMuFenXi = browser.find_element_by_class_name('anticon-eye')
        JieMuFenXi.click()
        sleep(1)
        DangRifenXi = browser.find_element_by_xpath('//*[@id="program-analysis$Menu"]/li[1]/a')
        DangRifenXi.click()
        sleep(1)
        JingZhengFenxi = browser.find_element_by_xpath('//*[@id="header___3l7sS"]/ul/li[2]/a')
        JingZhengFenxi.click()
        sleep(5)
        print('Get source Successfully...')
    except:
        print('Get source fail, please try again...')

def GetOptions(): # input is date string eg: '2020-01-01'
    # get BoFang Combobox
    Bofang = browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/main/main/section[2]/form/div[10]/div/div/span/div/div/span')
    Bofang.click()
    sleep(1)
    # select all
    all = browser.find_element_by_xpath('/html/body/div[2]/div/div/div/ul/li[3]')
    all.click()
    sleep(2)

def create_assist_date(datestart = None,dateend = None):
    if datestart is None:
        datestart = '2020-01-01'
    if dateend is None:
        dateend = datetime.datetime.now().strftime('%Y-%m-%d')

    datestart = datetime.datetime.strptime(datestart,'%Y-%m-%d')
    dateend = datetime.datetime.strptime(dateend,'%Y-%m-%d')
    date_list = []
    date_list.append(datestart.strftime('%Y-%m-%d'))
    while datestart < dateend:
        # add a day
        datestart += datetime.timedelta(days=+1)
        # convert datetime to string
        date_list.append(datestart.strftime('%Y-%m-%d'))
    return date_list

def renameFile(name):
    src = os.path.join(download_dir, '节目分析-当日分析-竞争分析.xlsx')
    dst = os.path.join(download_dir, '{}.xlsx'.format(name))
    os.rename(src, dst)


def Downloads(dateList):
    print('Downloading start...')
    for date in dateList:
        # get time inputbox
        timeinput = browser.find_element_by_class_name('ant-calendar-picker')
        timeinput.click()
        sleep(1)
        timecontnt = browser.find_element_by_class_name('ant-calendar-input ')
        timecontnt.send_keys(Keys.CONTROL, 'a')
        timecontnt.send_keys(date)
        sleep(1)
        timecontnt.send_keys(Keys.ENTER)
        sleep(1)
        # confirm the option
        confirm = browser.find_element_by_class_name('ant-btn-primary')
        confirm.click()
        sleep(3)
        # download based on the options
        download = browser.find_element_by_class_name('download___1hut5')
        download.click()
        sleep(5)
        print('The data on {} has been downloaded...'.format(date))
        # change the name if downloaded file
        renameFile(date)


if __name__ == '__main__':
    url = 'http://tv-zone.huan.tv/program-analysis/day-analysis/compete'
    dateList = create_assist_date(datestart='2020-01-01')
    browser.get(url)
    sleep(4)
    log_in()
    getsource()
    GetOptions()
    Downloads(dateList)
    browser.quit()
    print('The crawler is ended...')

