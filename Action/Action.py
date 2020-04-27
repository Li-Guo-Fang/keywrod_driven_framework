from selenium import webdriver
import time

from Utile.ParseConfig import *
from Conf.ProjVar import *
from Utile.ObjectMap import find_element,find_elements
from Utile.Excel import *


driver = ""

def is_xpath(exp):
    if ("//"in exp) or ("[" in exp) or ("@" in exp):
        return True
    return False

def get_element(driver,locator_exp):
    #print("当前定位的section和option:",locator_exp)
    section_name = locator_exp.split(",")[0]
    option_name = locator_exp.split(",")[1]
    element_locator = read_ini_file_option(
        PageElementLocator_file_path, section_name, option_name)
    element = find_element(
        driver, element_locator.split(">")[0], element_locator.split(">")[1])
    return element



def open_browser(browser_name):
    global driver
    if 'ie' in browser_name.lower():
        driver  = webdriver.Ie(executable_path="d:\\IEDriverServer")
    elif 'chrome' in browser_name.lower():
        driver  = webdriver.Chrome(executable_path="d:\\chromedriver")
    else:
        driver = webdriver.Firefox(executable_path="d:\\geckodriver")
    return driver

def visit(url):
    global driver
    driver.get(url)


def input(xpath_exp,content):
    global driver
    if is_xpath(xpath_exp):
        element = driver.find_element_by_xpath(xpath_exp)
        element.send_keys(content)
    else:
        element = get_element(driver,xpath_exp)
        element.send_keys(content)

def click(xpath_exp):
    global driver
    if is_xpath(xpath_exp):
        element = driver.find_element_by_xpath(xpath_exp)
        element.click()
    else:
        element = get_element(driver,xpath_exp)
        element.click()

def sleep(seconds):
    time.sleep(float(seconds))

def assert_word(expected_word):
    global driver
    assert expected_word in driver.page_source

def switch_to(xpath_exp):
    global driver
    if is_xpath(xpath_exp):
        driver.switch_to.frame(driver.find_element_by_xpath(xpath_exp))
    else:
        element = get_element(driver,xpath_exp)
        driver.switch_to.frame(element)

def switch_back():
    global driver
    driver.switch_to.default_content()

def quit():
    global driver
    driver.quit()

if __name__=="__main__":

    try:
        open_browser("chrom")
        visit("http://mail.126.com")
        click("126mail_login,login.loginlink")
        sleep(2)
        switch_to("126mail_login,login.frame")
        input("126mail_login,login.username","lgf18301991450")
        input("126mail_login,login.password","411527lgf")
        click("126mail_login,login.loginbutton")
        switch_back()
        sleep(2)
        assert_word("通讯录")
        click("126mail_home,homePage.addressbook")
        sleep(2)
        click("126mail_addContacts,addcontacts.ContactsBtn")
        sleep(1)
        input("126mail_addContacts,addcontacts.PersonName","lily")
        input("126mail_addContacts,addcontacts.PersonEmail","55w55@ewq.com")
        click("126mail_addContacts,addcontacts.star")
        input("126mail_addContacts,addcontacts.PersonMobile","11111111111")
        input("126mail_addContacts,addcontacts.PersonComment","work")
        click("126mail_addContacts,addContacts.save")


    except AssertionError as e:
        print("断言失败")
    except Exception as e:
        print("出现异常了",e)
    finally:
        quit()