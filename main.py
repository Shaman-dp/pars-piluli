from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from pathlib import Path
import requests
import os
from selenium.common.exceptions import NoSuchElementException
import xlwt
import httplib2
import time

BASE_URL = 'https://www.piluli.ru/category138_977_526/index.html'
#driver = webdriver.Chrome('D:\\Python\\driver\\chromedriver.exe')
driver = webdriver.Chrome()

#create new workbook
MAIN_WORKBOOK = xlwt.Workbook()

#create new page(sheet)
MAIN_SHEET = MAIN_WORKBOOK.add_sheet('New sheet')
#.write(row, column, data)
MAIN_SHEET.write(0, 0, 'Link_on_image')
MAIN_SHEET.write(0, 1, 'Name')
MAIN_SHEET.write(0, 2, 'Price')

BASE_SAVE_PATH = Path('./pars')
if not os.path.exists(BASE_SAVE_PATH):
	os.makedirs(BASE_SAVE_PATH)

FOR_IMAGE_PRODUCT = Path('./pars/image')
if not os.path.exists(FOR_IMAGE_PRODUCT):
	os.makedirs(FOR_IMAGE_PRODUCT)

driver.get(BASE_URL)

cookie = driver.find_element_by_class_name('cookie-close').find_element_by_class_name('piluli-54')
ActionChains(driver).move_to_element(cookie).click().perform()
print('cookie was killed')

#for name
number = 0

count_page = 1

#for __ xls
count_rows = 1

while True:

	print('page number: __________________________' + str(count_page))
	
	class_img = driver.find_elements_by_class_name('item-card')
			
	#parsing image
	for id_img in class_img:
		number += 1
			
		http = httplib2.Http('.cache')
		response, content = http.request(id_img.find_element_by_class_name('img._js-load-img').get_attribute('src'))

		#print in console for me
		name = id_img.find_element_by_class_name('item-name')
		#price = id_img.find_element_by_class_name('price-wrap')
		print(name.text)
		#print(price.text.split()[0])

		full_name_product = str(number)
		link_on_one_product_image = full_name_product + '.jpg'	

		#write a data
		MAIN_SHEET.write(count_rows, 0, full_name_product)#.write(row, column, data)
		MAIN_SHEET.write(count_rows, 1, name.text)
		#MAIN_SHEET.write(count_rows, 2, price.text.split()[0])

		MAIN_WORKBOOK.save('./pars/tranki.xls')
					
		count_rows += 1

		#write an image
		out = open(FOR_IMAGE_PRODUCT / link_on_one_product_image, "wb")
		out.write(content)
		out.close()
	
	#try find buttom
	time.sleep(3)
	if count_page > 2:
		if back_page == driver.find_element_by_class_name('piluli-36'):
			break
	if count_page > 1:
		back_page = driver.find_element_by_class_name('piluli-36')
	next_page = driver.find_element_by_class_name('piluli-36')
	ActionChains(driver).move_to_element(next_page).click().perform()

	count_page += 1
			
driver.quit()

print('Download [' + str(number) + '] files')