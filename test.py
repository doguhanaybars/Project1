
############################### Library ##########################################
import pandas as pd
import openpyxl
import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys
import os
import colorama
from colorama import Fore, Back, Style
import random
colorama.init()
############################# Date ###############################################
date=(datetime.datetime.now().strftime("%Y.%m.%d"))
hour=(datetime.datetime.now().strftime("%H.%M"))
############################# Case Klasorleri ####################################
    
class CreateDirectory():
    def ScreenShotDirectory(): 
        # Create directory
        dirName = f'ScreenShot'
        try:
            # Create target Directory
            os.mkdir(dirName)
            print(Fore.GREEN)
            print(dirName ,  "Klasoru olusturuldu ")
            print(Style.RESET_ALL)
        except FileExistsError:
            print(Fore.YELLOW)
            print(dirName ,  "Klasoru zaten mevcut")
            print(Style.RESET_ALL)
        # Create directory
        dirNameiki = f'ScreenShot/{date}'
        try:
            # Create target Directory
            os.mkdir(dirNameiki)
            print(Fore.GREEN)
            print(dirNameiki ,  "Klasoru olusturuldu ") 
            print(Style.RESET_ALL)
        except FileExistsError:
            print(Fore.YELLOW)
            print(dirNameiki ,  "Klasoru zaten mevcut")
            print(Style.RESET_ALL)
        # Create directory
        dirNameuc = f'ScreenShot/{date}/{hour}'
        try:
            # Create target Directory
            os.mkdir(dirNameuc)
            print(Fore.GREEN)
            print(dirNameuc ,  "Klasoru olusturuldu ") 
            print(Style.RESET_ALL)
        except FileExistsError:
            print(Fore.YELLOW)
            print(dirNameuc ,  "Klasoru zaten mevcut")
            print(Style.RESET_ALL)
    def LogsDirectory(): 
        # Create directory
        dirName = f'Logs'
        try:
            # Create target Directory
            os.mkdir(dirName)
            print(Fore.GREEN)
            print(dirName ,  "Klasoru olusturuldu ")
            print(Style.RESET_ALL)
        except FileExistsError:
            print(Fore.YELLOW)
            print(dirName ,  "Klasoru zaten mevcut")
            print(Style.RESET_ALL)
        # Create directory
        dirNameiki = f'Logs/{date}'
        try:
            # Create target Directory
            os.mkdir(dirNameiki)
            print(Fore.GREEN)
            print(dirNameiki ,  "Klasoru olusturuldu ") 
            print(Style.RESET_ALL)
        except FileExistsError:
            print(Fore.YELLOW)
            print(dirNameiki ,  "Klasoru zaten mevcut")
            print(Style.RESET_ALL)
        # Create directory
        dirNameuc = f'Logs/{date}/{hour}'
        try:
            # Create target Directory
            os.mkdir(dirNameuc)
            print(Fore.GREEN)
            print(dirNameuc ,  "Klasoru olusturuldu ") 
            print(Style.RESET_ALL)
        except FileExistsError:
            print(Fore.YELLOW)
            print(dirNameuc ,  "Klasoru zaten mevcut")
            print(Style.RESET_ALL)
    def ProductNameAndPrice():
        try:
            dirName = f'ProductNameAndPrice'
            # Create target Directory
            os.mkdir(dirName)
            print(Fore.GREEN)
            print(dirName ,  "Klasoru olusturuldu ") 
            print(Style.RESET_ALL)
        except FileExistsError:
            print(Fore.YELLOW)
            print(dirName ,  "Klasoru zaten mevcut")
            print(Style.RESET_ALL)

CreateDirectory.ScreenShotDirectory()
CreateDirectory.LogsDirectory()
CreateDirectory.ProductNameAndPrice()
############################# EXCEL ############################################
df=pd.read_excel('./Excel/Excel1.xlsx')
print(df.columns[0])
deger1=(df.columns[0])
df=pd.read_excel('./Excel/Excel2.xlsx')
print(df.columns[1])
deger2=(df.columns[1])
############################# SELENIUM ############################################
logs = open(f"./Logs/{date}/{hour}/{date}_{hour}.txt","w",encoding="utf-8") 
try:
    driver = webdriver.Edge(executable_path = '.\WebDriver\msedgedriver.exe')
    driver.set_window_size(1900,1000)
    driver.maximize_window()
    
    url =('https://www.beymen.com')
    driver.get(url)
    time.sleep(5)
    driver.save_screenshot(f"./ScreenShot/{date}/{hour}/anasayfa.png")
    time.sleep(1)
    print(Fore.GREEN)
    print('Anasayfa Goruntulendi')
    print(Style.RESET_ALL)
    logs.write(str('Anasayfa Goruntulendi'))

except:
    print(Fore.RED)
    print('HATA:Anasayfa Goruntulenmede Hata Alindi')
    print(Style.RESET_ALL)
    logs.write(str('HATA:Anasayfa Goruntulenmede Hata Alindi'))
try:
    WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"""/html/body/header/div/div/div[2]/div/div/div/input"""))).click()
    time.sleep(2)
    driver.find_element_by_xpath("/html/body/header/div/div/div[2]/div/div/div/input").send_keys(deger1)
    time.sleep(1)
    driver.find_element_by_xpath("/html/body/header/div/div/div[2]/div/div/div/input").send_keys(Keys.CONTROL, 'a')
    time.sleep(2)
    driver.find_element_by_xpath("/html/body/header/div/div/div[2]/div/div/div/input").send_keys(deger2)
    time.sleep(2)
    driver.save_screenshot(f"./ScreenShot/{date}/{hour}/arama1.png")
    time.sleep(1)
    driver.find_element_by_xpath("/html/body/header/div/div/div[2]/div/div/div/input").send_keys(Keys.ENTER)
    time.sleep(2)
    driver.save_screenshot(f"./ScreenShot/{date}/{hour}/arama2.png")
    time.sleep(1)
    print(Fore.GREEN)
    print('Arama Basarili')
    print(Style.RESET_ALL)
    logs.write(str('Arama Basarili'))
except:
    print(Fore.RED)
    print('HATA:Arama Basarisiz oldu')
    print(Style.RESET_ALL)
    logs.write(str('HATA:Arama Basarisiz oldu'))
logs.write(str('\n'))
try:
    random=random.randint(2,5)
    driver.find_element_by_xpath(f"//div[{random}]/div/div/div/a/div/div[2]").click()
    time.sleep(2)
    driver.save_screenshot(f"./ScreenShot/{date}/{hour}/kiyafet_secme.png")
    time.sleep(1)
    print(Fore.GREEN)
    print('Urun Secimi Basarili')
    print(Style.RESET_ALL)
    logs.write(str('Urun Secimi Basarili'))
except:
    print(Fore.RED)
    print('HATA:Urun Secimi Basarisiz oldu')
    print(Style.RESET_ALL)
    logs.write(str('HATA:Urun Secimi Basarisiz oldu'))
logs.write(str('\n'))
try:
    content = driver.find_elements_by_xpath("/html/body/div[3]/div[1]/div[1]/div[2]/div[2]/h1/span")
    time.sleep(2)
    for c in content:
       urunadi=c.text
       print('URUNUN ADI : ',urunadi)
    time.sleep(2)
    content = driver.find_elements_by_xpath("""//*[@id="priceNew"]""")
    for d in content:
       urunfiyat=d.text
       print('URUNUN FIYATI : ',urunfiyat)
    time.sleep(2)  
    dosya = open(f"./ProductNameAndPrice/{date}_{hour}.txt","w",encoding="utf-8")   
    write1 = 'URUNUN ADI :'+ str(urunadi)
    dosya.write(write1)
    dosya.write(str('\n'))
    write2 = 'URUNUN FIYATI :'+ str(urunfiyat)
    dosya.write(write2)
    time.sleep(2)
    print(Fore.GREEN)
    print('Bilgiler Basariyla Kaydedildi')
    print(Style.RESET_ALL)
    logs.write(str('Bilgiler Basariyla Kaydedildi'))
except:
    print(Fore.RED)
    print('HATA:Bilgiler Kaydedilirken Hata Alindi')
    print(Style.RESET_ALL)
    logs.write(str('HATA:Bilgiler Kaydedilirken Hata Alindi'))
logs.write(str('\n'))
try:
 
    driver.find_element_by_xpath(f"""//*[@id="sizes"]/div/span[1]""").click()
    time.sleep(2)
    driver.find_element_by_xpath(f"""//*[@id="addBasket"]""").click()
    time.sleep(5)
    driver.find_element_by_xpath(f"""/html/body/header/div/div/div[3]/div/a[3]/span""").click()
    time.sleep(2)
    driver.save_screenshot(f"./ScreenShot/{date}/{hour}/sepeteeklendi.png") 
    time.sleep(1)
    print(Fore.GREEN)
    print('Urun Sepete Eklendi')
    logs.write(str('Urun Sepete Eklendi'))
    print(Style.RESET_ALL)

except:
    print(Fore.RED)
    print('HATA:Urun Sepete Eklenirken Hata Alindi')
    print(Style.RESET_ALL)
    logs.write(str('HATA:Urun Sepete Eklenirken Hata Alindi'))
logs.write(str('\n'))

content = driver.find_elements_by_xpath("""/html/body/div[1]/div/div/div/div[1]/div[2]/div[1]/div[1]/div/div[1]/a/span""")
for c in content:
   urunadi2=c.text
   print('URUNUN ADI : ',urunadi2)
content = driver.find_elements_by_xpath("""/html/body/div[1]/div/div/div/div[1]/div[2]/div[1]/div[1]/div/div[1]/div[1]/div/div/span""")
for c in content:
   urunfiyat2=c.text
   print('URUNUN FIYATI : ',urunfiyat2)

if urunadi==urunadi2:
    print(Fore.GREEN)
    print('Urun Isimleri Birbirine Esit')
    print(Style.RESET_ALL)
    logs.write(str('Urun Isimleri Birbirine Esit'))
else:
    print(Fore.RED)
    print('Urun Isimleri Birbirine Esit Degil')
    print(Style.RESET_ALL)
    logs.write(str('HATA: Urun Isimleri Birbirine Esit Degil'))
logs.write(str('\n'))
if urunfiyat==urunfiyat2:
    print(Fore.GREEN)
    print('Urun Fiyatlari Birbirine Esit')
    print(Style.RESET_ALL)
    logs.write(str('Urun Fiyatlari Birbirine Esit'))
else:
    print(Fore.RED)
    print('Urun Fiyatlari Birbirine Esit Degil')
    print(Style.RESET_ALL)
    logs.write(str('HATA:Urun Fiyatlari Birbirine Esit Degil'))
logs.write(str('\n'))

try:
    driver.save_screenshot(f"./ScreenShot/{date}/{hour}/sepet.png")
    time.sleep(1)
    driver.find_element_by_xpath(f"""//*[@id="quantitySelect0-key-0"]""").click()
    time.sleep(2)
    time.sleep(2)
    driver.find_element_by_xpath(f"""//*[@id="quantitySelect0-key-0"]/option[2]""").click()
    time.sleep(2)
    driver.save_screenshot(f"./ScreenShot/{date}/{hour}/limit_arttirma.png")
    time.sleep(1)
    print(Fore.GREEN)
    print('Urun Sayisi 2 Olarak Arttirildi')
    print(Style.RESET_ALL)
    logs.write(str('Urun Sayisi 2 Olarak Arttirildi'))
except:
    print(Fore.RED)
    print('HATA:Urun Sayisi Arttirilirken Hata Alindi')
    print(Style.RESET_ALL)
    logs.write(str('HATA:Urun Sayisi Arttirilirken Hata Alindi'))
logs.write(str('\n'))

try:
    driver.find_element_by_xpath(f"""//*[@id="removeCartItemBtn0-key-0"]""").click()
    time.sleep(2)
    driver.save_screenshot(f"./ScreenShot/{date}/{hour}/sepet_temizlendi.png")
    time.sleep(1)
    print(Fore.GREEN)
    print('Sepet Temizlendi')
    print(Style.RESET_ALL)
    logs.write(str('Sepet Temizlendi'))
except:
    print(Fore.RED)
    print('HATA:Sepet Temizlenirken Hata Alindi')
    print(Style.RESET_ALL)
    logs.write(str('HATA:Sepet Temizlenirken Hata Alindi'))
logs.write(str('\n'))
driver.close()
print(Fore.GREEN)
print('TAMAMLANDI')
print(Style.RESET_ALL)
#input("ENTER tusuna basarak cikis yapabilirsiniz...")
