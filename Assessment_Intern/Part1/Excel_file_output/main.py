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
sheet['A1']='NAME'
sheet['B1']='PRICE'
sheet['C1']='RATING'
sheet['D1']='NO OF REVIEWS'
sheet['E1']='URL'
driver = webdriver.Chrome(ChromeDriverManager().install())
pgcount=1
j=2
for i in range(2,22):
    driver.get('https://www.amazon.in/s?k=bags&page='+str(pgcount)+'&crid=2M096C61O4MLT&qid=1667046635&sprefix=ba%2Caps%2C283&ref=sr_pg_'+str(pgcount))
    pgcount+=1
    items = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH, '//div[contains(@class, "s-result-item s-asin")]')))
    for item in items:
        try:
            name = item.find_element(by=By.XPATH, value='.//span[@class="a-size-medium a-color-base a-text-normal"]')
            sheet['A'+str(j)]=name.text
        except (NoSuchElementException,IndexError,AttributeError,UnicodeError) :
            sheet['A'+str(j)]='NOT FOUND'
            pass
        try:
            url=item.find_element(by=By.XPATH, value='.//a[@class="a-link-normal s-underline-text s-underline-link-text s-link-style a-text-normal"]').get_attribute('href')
            sheet['E'+str(j)]=url
        except (NoSuchElementException,IndexError,AttributeError,UnicodeError) :
            sheet['E'+str(j)]='NOT FOUND'
            pass 
        try:
            rating_box=item.find_elements(by=By.XPATH, value='.//div[@class="a-row a-size-small"]/span')
            ratings=rating_box[0].get_attribute('aria-label')
            sheet['C'+str(j)]=ratings
        except (NoSuchElementException,IndexError,AttributeError,UnicodeError) :
            sheet['C'+str(j)]='NOT FOUND'
            pass

        try:
            noofratings=rating_box[1].get_attribute('aria-label')
            sheet['D'+str(j)]=noofratings
        except (NoSuchElementException,IndexError,AttributeError,UnicodeError) :
            sheet['D'+str(j)]='NOT FOUND'
            pass

        try:
            price=item.find_element(by=By.XPATH, value='.//span[@class="a-price-whole"]')
            sheet['B'+str(j)]=price.text
        except (NoSuchElementException,IndexError,AttributeError,UnicodeError) :
            sheet['B'+str(j)]='NOT FOUND'
            pass
        j+=1
    if i<21:
        next=driver.find_element_by_css_selector('.s-pagination-next.s-pagination-separator')
        next.click()
path=os.getcwd()
os.mkdir(os.path.join(path,"Part1/Excel_file_output/Output"))
wb.save(os.path.join(path,"Part1/Excel_file_output/Output/Task1.xlsx"))
