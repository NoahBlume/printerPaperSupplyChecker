from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

#gets the google sheet that is to be edited
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('YourKey.json', scope)
gc = gspread.authorize(credentials)
wks = gc.open("Spreadsheet Name").sheet1

#A list of all of the printers that are to be checked (they are named after Peanuts characters)
printerDicts = [{'name': 'charliebrown', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'big'},
{'name': 'snoopy', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'big'},
{'name': 'woodstock', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'big'},
{'name': 'sally', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'small'},
{'name': 'spike', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'small'},
{'name': 'olaf', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'big'},
{'name': 'pigpen', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'small'},
{'name': 'schroeder', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'big'},
{'name': 'peppermintpatty', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'big'},
{'name': 'franklin', 'address': 'printeraddress.com', '8x11': 0, '8x14': 0, '11x17': 0, 'toner': False, 'size': 'big'}]

#sets up Chromedriver for Selenium
chromedriver = "/the/address/of/chromedriver/on/your/computer"
os.environ["webdriver.chrome.driver"] = chromedriver

driver = webdriver.Chrome(chromedriver)

#checks paper levels of each drawer in the printer
#toner checeking has not been implemented (I am waiting for one to actually become empty to see what the website will display)
def printerCheck(size):
	print("Now checking " + printer['name'])

	#goes to the address for the printer
	address = printer['address']
	driver.get(address)

	#logs into the site
	search_box = driver.find_element_by_name('username')
	search_box.send_keys('Your Username')
	search_box = driver.find_element_by_name('password')
	search_box.send_keys('Your Password')
	search_box.submit()

	#gets a list of all of the table rows; some of these will contain the paper levels
	trs = driver.find_elements(By.TAG_NAME, "tr")
	for aRow in trs:
		#checks levels of 11x17 paper
		if ("11x17" in aRow.text):
			#on the site, paper level is indicated by an image that indicates either full, 2/3 full, 1/3 full, or empty
			#this tries to find one of the images that indicates the paper is 1/3 full or empty
			oneThird = aRow.find_elements_by_xpath(".//img[@title='1-33 % ']")
			empty = aRow.find_elements_by_xpath(".//img[@title='None ']")
			#if 11x17 paper is needed, adds it to the amount in the printer's dictionary
			if len(oneThird) == 1:
				print('found an 11x17 1')
				printer['11x17'] += 1
			elif len(empty) == 1:
				print('found an 11x17 e')
				printer['11x17'] += 1
		#checks levels of 8x11 paper
		elif ("LTR" in aRow.text) or ("LTRR" in aRow.text):
			#this tries to find one of the images that indicates the paper is 2/3 full, 1/3 full, or empty
			twoThirds = aRow.find_elements_by_xpath(".//img[@title='33-66 % ']")
			oneThird = aRow.find_elements_by_xpath(".//img[@title='1-33 % ']")
			empty = aRow.find_elements_by_xpath(".//img[@title='None ']")
			#if 8x11 paper is needed, adds it to the amount in the printer's dictionary
			#checks if the printer has big drawers (these hold more paper and need different values added)
			#only drawers of number 1 and 2 can have big drawers (all the other drawers are small)
			if (size == 'big') and (("Drawer 1" in aRow.text) or ("Drawer 2" in aRow.text)):
				if len(twoThirds) == 1:
					printer['8x11'] += 1
				elif len(oneThird) == 1:
					printer['8x11'] += 2
				elif len(empty) == 1:
					printer['8x11'] += 3
			#otherwise the drawers are small, and need a different amount of paper added
			else:
				if len(oneThird) == 1:
					printer['8x11'] += 1
				elif len(empty) == 1:
					printer['8x11'] += 1
		#checks levels of 8x14 paper
		elif ("LGL" in aRow.text):
			#this tries to find one of the images that indicates the paper is 1/3 full or empty
			oneThird = aRow.find_elements_by_xpath(".//img[@title='1-33 % ']")
			empty = aRow.find_elements_by_xpath(".//img[@title='None ']")
			#if 8x14 paper is needed, adds it to the amount in the printer's dictionary
			if len(oneThird) == 1:
				printer['8x14'] += 1
			elif len(empty) == 1:
				printer['8x14'] += 1

	#updates the spreadsheet with the amount of paper needed for the printer
	strRow = str(currentRow)
	wks.update_acell('C' + strRow, str(printer['8x11']))
	wks.update_acell('D' + strRow, str(printer['8x14']))
	wks.update_acell('E' + strRow, str(printer['11x17']))
	return [printer['8x11'], printer['8x14'], printer['11x17']]


#the main stuff that runs

#the row in the spreadsheet that is being edited
currentRow = 2
#list that stores the total amount of each type of paper that's needed
totals = [0, 0, 0]
#goes through each printer and checks paper levels, updating the totals
for printer in printerDicts:
	toAdd = printerCheck(printer['size'])
	print("Needs: " + str(printer['8x11']) + " 8x11")
	print("Needs: " + str(printer['8x14']) + " 8x14")
	print("Needs: " + str(printer['11x17']) + " 11x17")
	i = 0
	while i < 3:
		totals[i] += toAdd[i]
		i += 1

	currentRow += 1

#closes the chrome window
driver.quit()

print("Totals:")
print("8x11 Needed: " + str(totals[0]))
print("8x14 Needed: " + str(totals[1]))
print("11x17 Needed: " + str(totals[2]))

#updates the spreadsheet to show the last time the supply levels were updated
lastUpdated = datetime.now().strftime("%m-%d-%Y %H:%M")
wks.update_acell('B14', lastUpdated)