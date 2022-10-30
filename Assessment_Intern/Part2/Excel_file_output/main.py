from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import openpyxl 
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import os

wb = openpyxl.Workbook() 
sheet = wb.active

sheet['A1']='DESCRIPTION'
sheet['B1']='ASIN'
sheet['C1']='PRODUCT DESCRIPTION'
sheet['D1']='MANUFACTURER'

driver = webdriver.Chrome(ChromeDriverManager().install())
path=os.getcwd()
rexcel=openpyxl.load_workbook(os.path.join(path+"/Part1/Excel_file_output/Output/Task1.xlsx"))
rsheet=rexcel.active

for r in range(2,203):
    asin='NOT FOUND'
    pdesc='NOT FOUND'
    manu='NOT FOUND'
    driver.get(rsheet['E'+str(r)].value)
    try:

        descli=driver.find_elements(by=By.XPATH, value='.//div[@id="feature-bullets"]/ul/li')
        desc=''
        for item in descli:
            txt=item.find_element(by=By.TAG_NAME, value='span')
            desc+=txt.text+' '
        sheet['A'+str(r)].value=desc
    except NoSuchElementException:
        sheet['A'+str(r)].value='NOT FOUND'
        pass
    try:
        asin=driver.find_element_by_id('olpLinkWidget_feature_div')
        sheet['B'+str(r)].value=asin.get_attribute('data-csa-c-asin')
    except NoSuchElementException:
        sheet['B'+str(r)].value='NOT FOUND'
        pass
    try:
        pdesc=driver.find_element(by=By.XPATH, value='.//div[@id="productDescription"]/p/span')
        sheet['C'+str(r)].value=pdesc.text
    except NoSuchElementException:
        sheet['C'+str(r)].value='NOT FOUND'
        pass
    try:
        ele=driver.find_element_by_id('bylineInfo').text
        sheet['D'+str(r)].value=ele[10:]
    except NoSuchElementException:
        sheet['D'+str(r)].value='NOT FOUND'
        pass
os.mkdir(os.path.join(path,"Part2/Excel_file_output/Output"))
wb.save(os.path.join(path+"/Part2/Excel_file_output/Output/Task2.xlsx"))