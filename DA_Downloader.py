###########################################################
# Day-Ahead Data downloader with selenium

import time
from selenium import webdriver
import datetime
from bs4 import BeautifulSoup
import pandas as pd

############################################################
# Insert the path to your Download folder here!
downloadlink = 'C:/Users/b7217/Downloads/'
# Insert the path to your chromedriver here!
driver = webdriver.Chrome('C:/webdrivers/chromedriver')
driver.maximize_window()
############################################################

#OKTE-------------------------------------------------------------------
driver.get('https://www.okte.sk/en/short-term-market/published-information/daily-stm-results/')
driver.find_element_by_xpath('/html/body/div[3]/div/section/form/div[2]/div/div[2]/div[2]/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[2]').click()
time.sleep(1)
button = driver.find_element_by_xpath('//*[@id="ext-gen64"]')
button.click()
time.sleep(1)
#OPCOM-------------------------------------------------------------------
driver.get('https://www.opcom.ro/pp/grafice_ip/raportPIPsiVolumTranzactionat.php?lang=en')
driver.find_element_by_xpath('//*[@id="menu_sari"]/option[text()="CSV (comma delimited)"]').click()
time.sleep(1)
#OTE-------------------------------------------------------------------
driver.get('https://www.ote-cr.cz/en/short-term-markets/electricity/day-ahead-market')
time.sleep(1)
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(1)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/article/div[2]/article/section[2]/div/div[2]/div[3]/div[2]/div/p[1]/a[1]').click()
time.sleep(1)
#HUPX-------------------------------------------------------------------
driver.get('https://hupx.hu/hu/piaci-adatok/dam/heti-adatok')
driver.execute_script("window.scrollTo(0, 300)") 
driver.find_element_by_xpath('/html/body/div[1]/div[3]/main/div[2]/div/div/div[1]/div[2]/div/a').click()
time.sleep(1)
#IBEX-------------------------------------------------------------------
driver.get('http://www.ibex.bg/en/market-data/dam/prices-and-volumes/')
time.sleep(1)
button3 = driver.find_element_by_xpath('//*[@id="download-button"]')
button3.click()
time.sleep(1)

#CROPEX-------------------------------------------------------------------
driver.get('https://www.cropex.hr/en/market-data/day-ahead-market/day-ahead-market-results.html')
soup = BeautifulSoup(driver.page_source, 'lxml')
DA = soup.find(id="cropex_dayahead").prettify()
#Convert to Dataframe
df = pd.read_html(DA)
#Convert to Excel
cropextable = df[0]
cropextable.to_excel(downloadlink+'CropexDA.xls')

#BSP-------------------------------------------------------------------
driver.get('https://www.bsp-southpool.com/trading-data.html')
time.sleep(1)
buttonbsp = driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/main/div[1]/div[2]/div/ul/li[1]/a')
buttonbsp.click()
time.sleep(1)

#EPEX -------------------------------------------------------------------
#EPEX DE link handler
epexlink = 'https://www.epexspot.com/en/market-data/dayaheadauction/auction-table/'
now = datetime.datetime.now() + datetime.timedelta(days=1)
date = now.strftime("%Y-%m-%d")
de = '/DE_LU'
fr = '/FR'
at = '/AT'
delink = epexlink+date+de
frlink = epexlink+date+fr
atlink = epexlink+date+at
#Selenium Working
driver.get(delink)
time.sleep(1)
driver.execute_script("window.scrollTo(0, 500)")
time.sleep(1)
#Passing Pagesource to Beutiful soup
soupepex = BeautifulSoup(driver.page_source, 'lxml')
epexDA = soupepex.find(id="tab_de_lu").prettify()
#Convert to Dataframe
edf = pd.read_html(epexDA)
#Convert to Excel
edf[2].to_excel(downloadlink+'EPEXDE.xls')

#Selenium Working FR
driver.get(frlink)
time.sleep(1)
driver.execute_script("window.scrollTo(0, 500)")
time.sleep(1)
soupepexde = BeautifulSoup(driver.page_source, 'lxml')
epexDAfr = soupepexde.find(id="tab_fr").prettify()
#Convert to Dataframe
edf = pd.read_html(epexDAfr)
#Convert to Excel
edf[2].to_excel(downloadlink+'EPEXFR.xls')
#Selenium Working AT
driver.get(atlink)
time.sleep(1)
driver.execute_script("window.scrollTo(0, 500)")
time.sleep(1)
soupepexfr = BeautifulSoup(driver.page_source, 'lxml')
epexDAat = soupepexfr.find(id="tab_at").prettify()
#Convert to Dataframe
edf = pd.read_html(epexDAat)
#Convert to Excel
edf[2].to_excel(downloadlink+'EPEXAT.xls')
#SEEPEX -------------------------------------------------------------------
driver.get('http://www.seepex-spot.com/sr/market-data/day-ahead-auction')
time.sleep(1)
driver.execute_script("window.scrollTo(0, 500)")
time.sleep(1)
#Passing Pagesource to Beutiful soup
soupepex = BeautifulSoup(driver.page_source, 'lxml')
srda = soupepex.find(class_="list hours responsive").prettify()
#Convert to Dataframe
srdf = pd.read_html(srda, decimal=',', thousands='.')
#Convert to Excel
srdf[0].to_excel(downloadlink+'seepex.xls')

driver.close()