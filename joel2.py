from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import xlsxwriter
import pandas as pd
from openpyxl import load_workbook

url = 'https://www.google.com/travel/hotels/Netherlands?g2lb=2502405%2C2502548%2C4208993%2C4254308%2C4258168%2C4260007%2C4270442%2C4274032%2C4285990%2C4291318%2C4301054%2C4305595%2C4308216%2C4308227%2C4313006%2C4314846%2C4315873%2C4317816%2C4317915%2C4324293%2C4326405%2C4328159%2C4329288%2C4270859%2C4284970%2C4291517%2C4292955%2C4316256%2C4333108&hl=en&gl=rs&un=1&q=b%26b%20holland&rp=ENSjnoWVlcfOpAEQnqjrgNzo2vDWARCG57_O0MSEvysQ_sHgo8_nzYf-ATgBQABIAg&ictx=1&ved=2ahUKEwi89ZmZqurmAhXkpYsKHa96Bu4QtccEegQIDBBO&hrf=CgYI0IYDEAAiA1JTRCoWCgcI5A8QARgGEgcI5A8QARgHGAEgArABAFgBaAGKASgKEgk0Z0b6YqVJQBEAEZAigSIOQBISCfZ2x_kMwEpAEYAISJFAcBlAmgENEgtOZXRoZXJsYW5kc6IBFwoIL20vMDU5ajISC05ldGhlcmxhbmRzqgExCgIIIRICCBUSAggNEgIIZxICCFsSAwiOARIDCJQCEgIILxICCFoSAwiSAhICCCAYAaoBDwoCCBISAwibARICCGgYAaoBDAoDCJ0BEgMIoAEYAaoBGgoCCBwSAghREgIIcxICCEcSAgg2EgIIKRgBqgEOCgIIJRICCHkSAgh6GAGqASIKAggREgIIKhICCEASAgg4EgIIVxICCAISAgh_EgIIKxgBqgEpCgIILhICCDsSAghWEgIIPRIDCIEBEgMIgwESAghLEgIIDBIDCIkBGAGqAQcKAwinARgAqgEWCgMIqQESAwirARIDCKoBEgMIrAEYAaoBDwoCCFASAwiEARICCE8YAaoBDgoCCDUSAggTEgIIMhgBkgECIAE&tcfs=EjEKCC9tLzA1OWoyEgtOZXRoZXJsYW5kcxoYCgoyMDIwLTAxLTA2EgoyMDIwLTAxLTA3UgA&ap=KigKEgk0Z0b6YqVJQBEAEZAigSIOQBISCfZ2x_kMwEpAEYAISJFAcBlAMAJaoAMKBgjQhgMQACIDUlNEKhYKBwjkDxABGAYSBwjkDxABGAcYASACsAEAWAFoAYoBKAoSCVBXoLN9r0lAEQARkCJBRA5AEhIJcc5KVqnJSkARgAhIkSCBGUCaAQ0SC05ldGhlcmxhbmRzogEXCggvbS8wNTlqMhILTmV0aGVybGFuZHOqATEKAgghEgIIFRICCA0SAghnEgIIWxIDCI4BEgMIlAISAggvEgIIWhIDCJICEgIIIBgBqgEPCgIIEhIDCJsBEgIIaBgBqgEMCgMInQESAwigARgBqgEaCgIIHBICCFESAghzEgIIRxICCDYSAggpGAGqAQ4KAgglEgIIeRICCHoYAaoBIgoCCBESAggqEgIIQBICCDgSAghXEgIIAhICCH8SAggrGAGqASkKAgguEgIIOxICCFYSAgg9EgMIgQESAwiDARICCEsSAggMEgMIiQEYAaoBBwoDCKcBGACqARYKAwipARIDCKsBEgMIqgESAwisARgBqgEPCgIIUBIDCIQBEgIITxgBqgEOCgIINRICCBMSAggyGAGSAQIgAYABAA'

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
#options.add_argument('--incognito')
options.add_argument('--headless')
driver = webdriver.Chrome("C:/python36/upwork/BandB_Joel/chromedriver", chrome_options=options)

def starting_url():
    driver.get(url)
    place = driver.find_elements_by_class_name('f1dFQe')
    length = len(place)
    return place, length

def get_data():

    name = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]')
    try:
        address = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[2]/span[5]/c-wiz/div/div/div/div/section[1]/div/span/span[2]')
        contact = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[2]/span[5]/c-wiz/div/div/div/div/section[2]/div/span/span[2]/span')
        website = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[2]/span[5]/c-wiz/div/div/div/div/section[2]/div/div/div/a')
    except Exception as e:
        try:
            address = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[2]/span[6]/c-wiz/div/div/div/div/section[1]/div/span/span[2]')
            contact = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[2]/span[6]/c-wiz/div/div/div/div/section[2]/div/span/span[2]/span')
            website = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[2]/span[6]/c-wiz/div/div/div/div/section[2]/div/div/div/a')
        except Exception as e:
            print('BAD XPATH')
            return {'Name': 'BAD XPATH', 'Address': 'BAD XPATH','Contact': 'BAD XPATH', 'Website': 'BAD XPATH'}
    #print(name.text)
    #print(address.text)
    #print(contact.text)
    #print(website.get_attribute('href'))
    return {'Name': name.text, 'Address': address.text,'Contact': contact.text, 'Website': website.get_attribute('href')}

def into_excel(row):
    #CRNA MAGIJA as cm
    #https://stackoverflow.com/questions/20219254/how-to-write-to-an-existing-excel-file-without-overwriting-data-using-pandas/47740262#47740262
    book = load_workbook('All_of_Ned.xlsx') #cm
    with pd.ExcelWriter('All_of_Ned.xlsx', engine = 'openpyxl', mode = 'a') as writer:
        writer.book = book #cm
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)#cm
        data = get_data()
        df = pd.DataFrame(data, columns=['Name', 'Address', 'Contact', 'Website'], index=[0])
        print(df)
        df.to_excel(writer, 'Main', header = False, index = False, startrow = row)
        #writer.save()
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

def switch_to_about():
    #driver.switch_to.window(driver.window_handles[1])
    try:
        driver.switch_to.window(driver.window_handles[1])
        about = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[1]/div[2]/div[1]/div[3]/div/div[5]/span')
        if about.text != 'About':
            about = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/div[1]/div[2]/div[1]/div[3]/div/div[6]/span')
            about.click()
        else:
            about.click()

    except Exception as e:
        return 'stop'

def first_next():
    time.sleep(2)
    button1 = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/c-wiz/div/div[1]/div[2]/div[4]/div/div[2]/c-wiz/div[6]/div/div/span')
    button1.click()
    time.sleep(2)
    place = driver.find_elements_by_class_name('f1dFQe')
    length = len(place)
    return place, length

def other_next():
    time.sleep(2)
    try:
        button = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/c-wiz/div/div[1]/div[2]/div[4]/div/div[2]/c-wiz/div[6]/div[2]/div/span/span')
        button.click()
        time.sleep(2)
    except Exception as e:
        try:
            button = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/c-wiz/div/div[1]/div[2]/div[4]/div/div[2]/c-wiz/div[6]/div[2]/div/div[2]')
            button.click()
            time.sleep(2)
        except Exception as e:
            button = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/c-wiz/div/div[1]/div[2]/div[4]/div/div[2]/c-wiz/div[6]/div[2]/div/span')
            button.click()
            time.sleep(2)

    place = driver.find_elements_by_class_name('f1dFQe')
    length = len(place)
    return place, length
