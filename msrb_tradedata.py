# -*- coding: utf-8 -*-
"""
Created on Thu Jul 18 13:25:12 2019

@author: Harsh Kava
"""

from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from fake_useragent import UserAgent
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import re,time,os,codecs,csv


#make browser
ua=UserAgent()
dcap = dict(DesiredCapabilities.PHANTOMJS)
dcap["phantomjs.page.settings.userAgent"] = (ua.random)
service_args=['--ssl-protocol=any','--ignore-ssl-errors=true']
driver = webdriver.Chrome('chromedriver.exe',desired_capabilities=dcap,service_args=service_args)


link='https://emma.msrb.org/TradeData/Search'
timeout = 60

    
#visit the link
driver.get(link)

but1 = driver.find_element_by_id('ctl00_mainContentArea_disclaimerContent_yesButton')
but1.click()

time.sleep(2)

data =[]  #all rows conatining data


#Months for which Reports needs to be generated
reportDate = {}
reportDate['Jan'] = ['01/01/2019','01/31/2019']
reportDate['Feb'] = ['02/01/2019','02/28/2019']
reportDate['Mar'] = ['03/01/2019','03/31/2019']
reportDate['Apr'] = ['04/01/2019','04/30/2019']
reportDate['May'] = ['05/01/2019','05/31/2019']
reportDate['Jun'] = ['06/01/2019','06/30/2019']
reportDate['Jul'] = ['07/01/2019','07/31/2019']


#Processing for every Month
for key, value in reportDate.items():
    #print(value[0])  #start date
    #print(value[1])  #end Date
    print(key)
    
    tradeDateFrom = driver.find_element_by_id('tradeDateFrom')
    tradeDateFrom.clear()
    tradeDateFrom.send_keys(value[0])      # send date in MM/DD/YYY format. ex:07/17/2019
    
    tradeDateTo = driver.find_element_by_id('tradeDateTo')
    tradeDateTo.clear()
    tradeDateTo.send_keys(value[1])

    #list of CUSIPS to be processed
    cusiplist =['63165TYM1','68441MAX3','544525SC3','93974DHC5','64990FHW7' \
                ,'59333RGV0','915183C47','874476HV9','59261ANA1','63165TWJ0','68607DPS8' \
                ,'64971QCH4','708399AB6','71883RLT8','691582FP0','576544S92','64971XGQ5' \
                ,'74526QUZ3','74526QWP3','74526QXU1','73358WZT4','451913AH0','64966L2M2' \
                ,'957297EM2','649717TU9','64971W7Q7','41423PAE7','91514AGB5','60412AMQ3' \
                ,'650014TN3','60412AML4','938782DP1','89602NE91']
    
    for c in cusiplist:
        
        print('Processing for Cusip::',c)
        
        #enter cusip value in CUSIP field
        cinput = driver.find_element_by_id('cusip')
        cinput.clear()
        cinput.send_keys(c)
        
        #click the search button
        runSearch = driver.find_element_by_id('searchButton')
        runSearch.click()
        

        time.sleep(2)
        #wait for the results to be populated
        
        try:
            table = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'lvSearchResults')))
            #table = driver.find_element_by_id('lvSearchResults')
            #nextPage = driver.find_element_by_id('lvSearchResults_next')
            while True:
                nextPage = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'lvSearchResults_next')))
                        
                #data in table
                rows = table.find_elements_by_tag_name('tr')
                for tr in rows:
                    tds = tr.find_elements_by_tag_name('td')
                    if tds:
                        row_data = []
                        idx = 0
                        for td in tds:
                            if idx == 1 :
                                row_data.append(c)
                                idx += 1
                            else:
                                row_data.append(td.text)
                                idx += 1
                        if(row_data[0] != ''):
                            data.append(row_data)

                #print(file_header) 
                #print(data)

                #Click Next Button until it is active
                if 'disabled' in nextPage.get_attribute('class'):
                    break;
                nextPage.click()
                time.sleep(3)
            
        except Exception as e:
            continue
        

#write csv from table
try:
    with open('Trade_Results.csv', mode='w', newline='') as trade_file:
        writer = csv.writer(trade_file, delimiter=',')
    
        writer.writerow(['Trade Date/Time','CUSIP *','Security Description *','Coupon(%)','MaturityDate','Price (%)','Yield (%)','Calculation/ Date & Price (%)','TradeAmount($)','SpecialCondition'])
        for d in data:
            writer.writerow(d)
            
        print('Trade_Results.csv file created.')
except Exception as e:
    print(e)
    print('Could not create csv')
    