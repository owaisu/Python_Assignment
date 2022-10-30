from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import csv
import os

head=['NAME','PRICE','RATING','NO OF REVIEWS','URL']
rows=[]
driver = webdriver.Chrome(ChromeDriverManager().install())
pgcount=1

for i in range(2,22):
    driver.get('https://www.amazon.in/s?k=bags&page='+str(pgcount)+'&crid=2M096C61O4MLT&qid=1667046635&sprefix=ba%2Caps%2C283&ref=sr_pg_'+str(pgcount))
    pgcount+=1
    items = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH, '//div[contains(@class, "s-result-item s-asin")]')))
    for item in items:
        name='Not Found'
        price='Not Found'
        ratings='Not Found'
        noofratings='Not Found'
        url='Not Found'
        try:
            name = item.find_element(by=By.XPATH, value='.//span[@class="a-size-medium a-color-base a-text-normal"]')
            url=item.find_element(by=By.XPATH, value='.//a[@class="a-link-normal s-underline-text s-underline-link-text s-link-style a-text-normal"]').get_attribute('href')
            rating_box=item.find_elements(by=By.XPATH, value='.//div[@class="a-row a-size-small"]/span')
            ratings=rating_box[0].get_attribute('aria-label')
            noofratings=rating_box[1].get_attribute('aria-label')
            price=item.find_element(by=By.XPATH, value='.//span[@class="a-price-whole"]')
            rows.append([name.text,price.text,ratings,noofratings,url])
        except (NoSuchElementException,IndexError,AttributeError,UnicodeError) :
            pass
    if i<21:
        next=driver.find_element(by=By.CSS_SELECTOR, value='.s-pagination-next.s-pagination-separator')
        next.click()


path=os.getcwd()
os.mkdir(os.path.join(path,"Part1/CSV_file_output/Output"))

try:
    with open(os.path.join(path,"Part1/CSV_file_output/Output/Task1.csv"),'w') as csf:
        csw=csv.writer(csf,delimiter=',')
        csw.writerow(head)
        csw.writerows(rows)
except UnicodeError:
    pass
