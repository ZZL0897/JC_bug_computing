from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook
import os
import time

if __name__ == '__main__':
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(r'D:\code\bug\chromedriver.exe', chrome_options=options)
    driver.get('https://www.pk6636.com/main.aspx')
    time.sleep(40)
    js = 'for (var i=0;i<15;i++){document.getElementsByClassName("clip")[i].style.display="block";}'

    driver.execute_script(js)

    treelink = driver.find_elements_by_class_name('treelink')
    print(treelink)
    for t in treelink:
        driver.execute_script("arguments[0].click();", t)
        print(t.text)





