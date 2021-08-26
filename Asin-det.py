from datetime import datetime
import eel
import time
import  requests , zipfile, io
import threading
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import TimeoutException
import os, sys
from selenium.webdriver.common.touch_actions import TouchActions
import tkinter as tk
import random
from tkinter import filedialog
import pandas as pd
import gspread
import html
import bs4
import html.parser
import lxml
import ctypes 
from bs4 import BeautifulSoup
import re
from oauth2client.service_account import ServiceAccountCredentials
root = tk.Tk()
root.withdraw()
import easygui
global files2
global star2
global en2
global thr1
global thr2
global thr3
global thr4

global getval
global openn
global signn
openn=False
getval=False
global atlist
global unavail
signn={'ASIN':{}}
unavail={'ASIN':[],'SKU':[]}
atlist=[]
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name(
         'login.json', scope) # Your json file here
global gc
gc = gspread.authorize(credentials)

sheet = gc.open("All information").worksheet('information')
def openmo():
    global openn
    openn=True
def closee():
    global openn
    openn=False
def asin1():
    global openn
    print("got here")
    global files2
    global star2
    global en2
    files=files2
    global lk
    star=int(star2)
    global lo
    en=int(en2)
    global assure
    #print(files)
    global gk
    #star=star
    global assure
    assure=False
    j=star
    options = webdriver.ChromeOptions() 

    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data2")
    #driver= webdriver.Chrome(executable_path=r"driver2/chromedriver.exe",options=options)
    #driver.get("https:www.amazon.com/dp/B01N9QKU2S")
    #time.sleep(3)
    #print(driver.find_element_by_xpath("//*[@id='priceBadging_feature_div']/i").get_attribute("aria-label"))
    #driver.quit()
    y="unavailable"
    x="unavailable"
    xy=[1,1,1,1,1,1,1,1,1,1,1]
    #options.add_argument('--headless')
    if openn==True:
        pass
    else:
        options.add_argument('--headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    wb=xw.Book("C:\\Users\\Public\\"+str(files))
    sht1=wb.sheets['New File']
    #assure=False
    #star=2
    sht1['A1'].value='Asins'
    sht1['C1'].value='Links'
    sht1['D1'].value='Title'
    sht1['E1'].value='Brands'
    sht1['F1'].value='Price'
    sht1['G1'].value='Ratings'
    sht1['H1'].value='Prime/Any Seller Avaialable'
    sht1['I1'].value='Quantity'
    sht1['J1'].value='weight in ounces'
    sht1['K1'].value='Dimensions'
    sht1['L1'].value='Launch date'
    sht1['M1'].value='Price1'
    sht1['N1'].value='Seller1'
    sht1['O1'].value='Price2'
    sht1['P1'].value='Seller2'
    sht1['Q1'].value='Price3'
    sht1['R1'].value='Seller3'
    sht1['S1'].value='Price4'
    sht1['T1'].value='Seller4'
    sht1['U1'].value='Price5'
    sht1['V1'].value='Seller5'
    sht1['W1'].value='Price6'
    sht1['X1'].value='Seller6'
    sht1['Y1'].value='Price7'
    sht1['Z1'].value='Seller7'
    sht1['AA1'].value='Price8'
    sht1['BB1'].value='Seller8'
    safe=star
    global getval
    global signn
    getval  =getval
    getval = False
    hal=1
    for iml in range(1,100):

        if getval==True:

            print("getvalll")
            break
        
        #options = webdriver.ChromeOptions()
        try:
        
            #options= webdriver.ChromeOptions()

            try:
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            except Exception:
                time.sleep(5)
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            else:
                pass
            
            for j in range(star,en+1,4):

                if hal%500==0:
                    time.sleep(300)
                hal=hal+1
                glis=[]
                glance=[]
                safe=j
                drive_call=False
                ydrive_call=False
                Zdrive_call=False
                try:
                    col=sht1['A'+str(j)].value
                except Exception:
                    print("XL")
                    #time.sleep(10)
                    print(expected)
                else:
                    pass
                print("driver3")
                print(col)
                try:

                    
                    if assure==True:
                        break
                    else:
                        for ite in range(1,30):
                            try:
                                driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                time.sleep(random.randint(3,7))
                                #time.sleep(2000)
                                
                                drive_call=True
                                time.sleep(1)
                            except (WebDriverException,TimeoutException) as w:
                                print(w)
                                time.sleep(3)
                                driver.refresh()
                            else:
                                break
                        print()

                        if drive_call==True:
                            pass
                        else:
                            
                            for ite in range(1,90):
                                time.sleep(10)
                                try:
                                    
                                    driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                    time.sleep(random.randint(3,7))
                                    #drive_call=True
                                    time.sleep(1)
                                except (WebDriverException,TimeoutException) as w:
                                    time.sleep(3)
                                    driver.refresh()
                                else:
                                    break
                            print()
                        
                    get_title = driver.title
                    asi_call=False
                    for prod in range(1,5):
                        try:
                            a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                            print('1')
                        except Exception as pro:
                            if not driver.title:
                                empty=True
                                break
                            elif "Page Not Found" in get_title:
                                break

                            else:
                                #print(y)
                                print(pro)
                                time.sleep(5)
                                a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                        else:
                            #print(a)
                            ASIN=False
                            try:
                                html= str(driver.page_source)
                                soup = BeautifulSoup(html,features="lxml")
                                finalhtml=soup.text.replace('\n', ' ')
                                print('1')
                            except Exception as pro:
                                pass
                            else:
                                #print(a)
                                if str(col) in finalhtml:
                                    print('art')
                                    print(hammer)
                                elif 'ASIN' in finalhtml:
                                    pass
                                else:
                                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                                    try:
                                        xi=driver.find_elements_by_tag_name('tbody')
                                        print('8')
                                        print(xi)
                                    except Exception:
                                        pass
                                    else:
                                        for jop in xi:
                                            
                                            if 'ASIN' in jop.get_attribute('innerHTML'):
                                                ASIN=True
                                                #print(jop.get_attribute('innerHTML'))
                                                if str(col) in jop.get_attribute('innerHTML'):
                                                    print(hammer2)
                                                    print('1')
                                                else:
                                                    print('1')
                                                    pass

                                        if ASIN==True:
                                            print('3')
                                            #pass
                                        else:
                                            print(hammer3)
                    
                except Exception as km:
                    #print('km')
                    print('2')
                    print(km)

                    #action = ActionChains(driver)
                    #firstLevelMenu = driver.find_element_by_id("acrPopover")
                    #action.move_to_element(firstLevelMenu).perform()
                    #time.sleep(5)
                    
                        
                    u="https://www.amazon.com/dp/"+str(col)
                    #u=' '
                    #uv="=HYPERLINK("+"https://www.amazon.com/dp/"+str(col)+, "product link")
                    glis.append(u)
                    #print(glis)
                    try:
                        a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(a)
                        glis.append(a)



                    try:
                        b=driver.find_element_by_id("bylineInfo_feature_div").text
                    except Exception:
                        print(y)
                        glis.append(y)
                    else:
                        #print(b)
                        glis.append(b)
                    print('ok')



                    try:
                        c=driver.find_element_by_xpath("//*[@id='priceblock_ourprice']").text

                    except Exception:
                        try:
                            z=driver.find_element_by_xpath("//*[@id='priceblock_businessprice']").text
                            
                        except Exception:
                            try:
                                q=driver.find_element_by_xpath("//*[@id='priceblock_saleprice']").text
                            except Exception:
                                try:
                                    p=driver.find_element_by_xpath("//*[@id='priceblock_dealprice']").text 
                                except Exception:
                                    try:
                                        zp=driver.find_element_by_xpath("//*[@id='priceblock_saleprice_lbl']").text
                                    except Exception:
                                        #print("none")
                                        glis.append("none")
                                    else:
                                        #print(p)
                                        glis.append(zp)
                                else:
                                    glis.append(p)
                            else:
                                #print(q)
                                glis.append(q)
                        else:
                            glis.append(z)
                    else:
                        #print(c)
                        glis.append(c)

                    #glis.append('')
                    #glis.append('')
                    try:
                        u=driver.find_element_by_id('acrPopover').get_attribute('title')
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(u)
                        glis.append(u)



                    try:
                        #ama=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-0']/span[2]/span").get_attribute('innerHTML')
                        #ama2=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-1']/span[2]/span").get_attribute('innerHTML')
                        ama=driver.find_element_by_css_selector("#tabular-buybox-truncate-0 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        ama2=driver.find_element_by_css_selector("#tabular-buybox-truncate-1 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        soup1 = BeautifulSoup(ama,features="lxml").text
                        soup2 = BeautifulSoup(ama2,features="lxml").text
                        
                        #print(bdy)
                        glis.append('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        print('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        
                        
                    except Exception:
                        glis.append("not available")
                    else:
                        pass


                    try:
                        el=driver.find_element_by_id("availability").text
                        try:
                            for x in range(1,10):
                                el=el.replace("\n", " ")
                        except Exception:
                            pass
                    except Exception:
                        print(y)
                        glis.append(y)
                        el=y
                    else:
                        #print(el)
                        if el:
                            glis.append(el)
                        else:
                            glis.append(y)
                            el=y


                    got=False
                    got2=False
                    fina='none'
                    dim=False
                    rep=re.split('\s+', finalhtml)
                    for t in rep:
                        if 'ounces' in t.lower():
                            try:
                                e=rep.index('ounces')
                                print(rep[e-1] +' '+'ounces')
                                glis.append(0.5)
                                dim=True
                                break
                            except Exception as r:
                                print(r)
                        else:
                            try:
                                if 'pounds' in t.lower():
                                    print(t)
                                    e=rep.index('pounds')
                                    print(rep[e-1]+' '+'pounds')
                                    glis.append(rep[e-1])
                                    dim=True
                                    break
                                else:
                                    pass
                            except Exception as p:
                                pass
                    if dim==True:
                        pass
                    else:
                        glis.append('not available')
                    inch=False    
                    date_avail=False
                    for ter in rep:
                        if 'inches' in ter:
                            rep2=rep.index(ter)
                            if rep[rep2-2]=="x":
                                pu=rep[rep2-5]+rep[rep2-4]+rep[rep2-3]+rep[rep2-2]+rep[rep2-1]+' '+rep[rep2]
                                glis.append(pu)
                                inch=True
                                break
                    for ter in rep:
                        if inch==True:
                            break
                        if 'Dimensions' in ter:
                            rep2=rep.index(ter)
                            print(rep[rep2-2])
                            if rep[rep2+2]=="x":
                                pu=rep[rep2+1]+rep[rep2+2]+rep[rep2+3]+rep[rep2+4]+rep[rep2+5]
                                glis.append(pu)
                                inch=True
                                break
                    if inch==False:
                        glis.append(y)
                    for tor in rep:    
                        if 'Date' in tor:

                            rep2=rep.index(tor)
                            if "First" in rep[rep2+1]:
                                print(rep[rep2-1])
                                pu2=rep[rep2+3]+' '+rep[rep2+4]+' '+rep[rep2+5]
                                if ':' in pu2:
                                    break
                                glis.append(pu2)
                                date_avail=True
                                break

                    if ':' in pu2:
                        for tor in rep:    
                            if 'Date' in tor:

                                rep2=rep.index(tor)
                                if "First" in rep[rep2+1]:
                                    print(rep[rep2-1])
                                    pu2=rep[rep2+6]+' '+rep[rep2+7]+' '+rep[rep2+8]
                                    glis.append(pu2)
                                    date_avail=True
                                    break
                    if date_avail==False:
                        glis.append(y)

                    put=False
                    put2=False
                    dr=[]
                    for jt in range(1,3):
                        
                        try:
                            driver.execute_script("arguments[0].setAttribute('class', 'a-section a-spacing-none a-spacing-top-small a-padding-none')",driver.find_element_by_xpath('//*[@id="aod-pinned-offer-additional-content"]'))    
                            #driver.find_element_by_id("aod-pinned-offer-show-more-link").click()
                            time.sleep(2)
                            dr=driver.find_element_by_id('aod-pinned-offer').text.split('\n')
                            print(dr)
                            driver.save_screenshot('screenie.png')
                            break
                        except Exception:
                            time.sleep(3)
                            driver.save_screenshot('screenie.png')
                    try:
                        #print(driver.find_element_by_id("aod-pinned-offer-additional-content").text)
                    
                        print(dr)
                        try:
                            x=dr.index('New')
                        except Exception:
                            x=dr.index('New - Business Price')
                        print(dr[x+1]+'.'+dr[x+2])
                        glis.append(str(dr[x+1]+'.'+dr[x+2]))
                        yr=dr.index('Ships from')
                        print(' '.join(dr[yr:]))
                        glis.append(' '.join(dr[yr:]))
                        put2=True
                    except Exception:
                        pass


                    try:
                        driver.execute_script("document.getElementById('aod-filter-list').style.display = 'block';")
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("primeEligible"))
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("new"))
                        print(driver.find_element_by_id('aod-filter-list').text)
                        time.sleep(2)
                        fo=driver.find_elements_by_id("aod-offer")
                        iv=[]
                        for io in fo:
                            ok=io.text
                            iv.append(ok.split('\n'))
                        yo=0
                        print(iv)
                        for a in iv:
                            x=iv[yo].index('New')
                            print(iv[yo][x+1]+'.'+iv[yo][x+2])
                            glis.append(str(iv[yo][x+1]+'.'+iv[yo][x+2]))
                            yr=iv[yo].index('Ships from')
                            print(' '.join(iv[yo][yr:]))
                            glis.append(' '.join(iv[yo][yr:]))
                            yo=yo+1
                            put=True
                    except Exception:
                        glis.append('NoNe')
                        glis.append('NoNEE')

                    m=1

                    glance=[]
                    
                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,90):
                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1
                        print()
                    except Exception as klo:
                        print(klo)
                    try:
                        if put2==True:
                            if 'stock soon' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'months'in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'weeks' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'out of stock' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif '1 left' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'unavailable' in el.lower():
                                #signn['ASIN'][j]=col
                                sht1['H'+str(j)].value='NO'
                            elif 'released' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'preorder' in  el.lower():
                                sht1['H'+str(j)].value='NO'
                            else:
                                sht1['H'+str(j)].value='YES'
                        elif not iv:
                            sht1['H'+str(j)].value='NO'
                        else:
                            pass
                        if put==False:
                            if 'NoNe' in glis[10]:
                                sht1['H'+str(j)].value='NO'
                        else:
                            sht1['H'+str(j)].value='YES'

                        
                        
                    except Exception:
                        pass
                    if 'unavailable' in el.lower():
                        signn['ASIN'][j]=col
                        

                    #nil=nil+1
                    #driver.quit
                    print(glis[3])
                    glis=[]
                    glance=[]
                else:
                    glance=[]
                    m=1
                    urr="https://www.amazon.com/dp/"+str(col)
                    #print(col)
                    glis=[]
                    
                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                    unavail['ASIN'].append(col)
                    unavail['SKU'].append(sht1['B'+str(j)].value)
                    #sheet.insert_row(glis,2)
                    

                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,67+len(xy)):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as kp:
                        print(kp)
                    sht1['H'+str(j)].value='NO'
                    glance=[]
                    
                    glis=[]
                #nil=nil+1
                try:
                    driver.find_element_by_xpath("//*[@id='nav-logo']/a/span")
                except Exception as ou:
                    print(ou)
                    
                    try:
                        driver.find_element_by_xpath("//*[@id='c']")
                    except Exception as og:
                        print(og)
                        if col=='None':
                            pass
                        elif empty==True:
                            pass
                        else:
                            print(ogg)
                    else:
                        try:
                            for l in range(67,67+len(xy)):

                                sht1[chr(l)+str(j)].value='unavailable'
                                m=m+1

                            print()
                        except Exception as kp:
                            print(kp)
                    
                else:
                    pass
                star=star+1

                
                #glance_called=False


                



            print()

        except Exception as hee:

            print('error here')
            star=safe
            try:
                
                driver.quit()
            except Exception:
                pass
            print(hee)
            print("xyxyxy")
            now = datetime.now()
    
            current_time = now.strftime("%I:%M:%p")
    
            
            pop13=["simple det error@ thread1",'index '+str(j),current_time]
            try:
                sheet_log = gc.open("All information").worksheet(gk)
                sheet_log.insert_row(pop13,4)
                #print('look')
            except Exception:
                pass
        else:
            
            
            try:
                
                driver.quit()
                
            except Exception:
                print('interrupted')
            else:
                print('completed')

                break
        finally:
            atlist.append(1)
            if len(atlist)==4:
                print('signn')
                print(signn)
                sign6()
                print('unavail')
                print(unavail)
                df=pd.DataFrame.from_dict(unavail,orient='index').transpose()
                print(df)
                if unavail['ASIN']:
                    #self.spapii()
                    pass
                else:
                    pass
            print("It's done finally")

def asin2():
    global openn
    print("got here2")
    global files2
    global star2
    global en2
    files=files2
    global lk
    star=int(star2)
    global lo
    en=int(en2)
    global assure
    #print(files)
    global gk
    #star=star
    global assure
    assure=False
    j=star
    options = webdriver.ChromeOptions() 

    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data3")
    #options.add_argument('--headless')
    if openn==True:
        pass
    else:
        options.add_argument('--headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    #driver= webdriver.Chrome(executable_path=r"driver2/chromedriver.exe",options=options)
    #driver.get("https:www.amazon.com/dp/B01N9QKU2S")
    #time.sleep(3)
    #print(driver.find_element_by_xpath("//*[@id='priceBadging_feature_div']/i").get_attribute("aria-label"))
    #driver.quit()
    y="unavailable"
    x="unavailable"
    xy=[1,1,1,1,1,1,1,1,1,1,1]
            
    wb=xw.Book("C:\\Users\\Public\\"+str(files))
    sht1=wb.sheets['New File']
    #assure=False
    #star=2
    safe=star
    global signn
    global getval
    getval=getval
    getval=False
    hal=1
    for iml in range(1,100):

        if getval==True:

            print("getvalll")
            
        
            break
        
        #options = webdriver.ChromeOptions()
        try:
            #options= webdriver.ChromeOptions()

            try:
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            except Exception:
                time.sleep(5)
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            else:
                pass
            
            for j in range(star+1,en+1,4):
                if hal%500==0:

                    time.sleep(300)
                hal=hal+1
                glis=[]
                glance=[]
                safe=j-1
                drive_call=False
                ydrive_call=False
                Zdrive_call=False
                try:
                    col=sht1['A'+str(j)].value
                except Exception:
                    print("XL")
                    #time.sleep(10)
                    print(expected)
                else:
                    pass
                print("driver3")
                print(col)
                try:

                    
                    if assure==True:
                        break
                    else:
                        for ite in range(1,30):
                            try:
                                driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                time.sleep(random.randint(3,7))
                                #time.sleep(2000)
                                
                                drive_call=True
                                time.sleep(1)
                            except (WebDriverException,TimeoutException) as w:
                                print(w)
                                time.sleep(3)
                                driver.refresh()
                            else:
                                break
                        print()

                        if drive_call==True:
                            pass
                        else:
                            
                            for ite in range(1,90):
                                time.sleep(10)
                                try:
                                    
                                    driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                    time.sleep(random.randint(3,7))
                                    #drive_call=True
                                    time.sleep(1)
                                except (WebDriverException,TimeoutException) as w:
                                    time.sleep(3)
                                    driver.refresh()
                                else:
                                    break
                            print()
                        
                    get_title = driver.title
                    asi_call=False
                    for prod in range(1,5):

                        try:
                            a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                            
                        except Exception as pro:
                            if not driver.title:
                                empty=True
                                break
                            elif "Page Not Found" in get_title:
                                break

                            else:
                                #print(y)
                                print(pro)
                                time.sleep(5)
                                a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                        else:
                            #print(a)
                            ASIN=False
                            try:
                                html= str(driver.page_source)
                                soup = BeautifulSoup(html,features="lxml")
                                finalhtml=soup.text.replace('\n', ' ')
                            except Exception as pro:
                                pass
                            else:
                                #print(a)
                                if str(col) in finalhtml:
                                    print(hammer)
                                elif 'ASIN' in finalhtml:
                                    pass
                                else:
                                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                                    try:
                                        xi=driver.find_elements_by_tag_name('tbody')
                                        
                                    except Exception:
                                        pass
                                    else:
                                        for jop in xi:
                                            
                                            if 'ASIN' in jop.get_attribute('innerHTML'):
                                                ASIN=True
                                                #print(jop.get_attribute('innerHTML'))
                                                if str(col) in jop.get_attribute('innerHTML'):
                                                    print(hammer2)
                                                    
                                                else:
                                                    pass

                                        if ASIN==True:
                                            pass
                                        else:
                                            print(hammer3)
                    
                except Exception as km:
                    #print('km')
                    print(km)

                    #action = ActionChains(driver)
                    #firstLevelMenu = driver.find_element_by_id("acrPopover")
                    #action.move_to_element(firstLevelMenu).perform()
                    #time.sleep(5)
                    
                        
                    u="https://www.amazon.com/dp/"+str(col)
                    #u=' '
                    #uv="=HYPERLINK("+"https://www.amazon.com/dp/"+str(col)+, "product link")
                    glis.append(u)
                    #print(glis)
                    try:
                        a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(a)
                        glis.append(a)



                    try:
                        b=driver.find_element_by_id("bylineInfo_feature_div").text
                    except Exception:
                        print(y)
                        glis.append(y)
                    else:
                        #print(b)
                        glis.append(b)
                    print('ok')



                    try:
                        c=driver.find_element_by_xpath("//*[@id='priceblock_ourprice']").text

                    except Exception:
                        try:
                            z=driver.find_element_by_xpath("//*[@id='priceblock_businessprice']").text
                            
                        except Exception:
                            try:
                                q=driver.find_element_by_xpath("//*[@id='priceblock_saleprice']").text
                            except Exception:
                                try:
                                    p=driver.find_element_by_xpath("//*[@id='priceblock_dealprice']").text 
                                except Exception:
                                    try:
                                        zp=driver.find_element_by_xpath("//*[@id='priceblock_saleprice_lbl']").text
                                    except Exception:
                                        #print("none")
                                        glis.append("none")
                                    else:
                                        #print(p)
                                        glis.append(zp)
                                else:
                                    glis.append(p)
                            else:
                                #print(q)
                                glis.append(q)
                        else:
                            glis.append(z)
                    else:
                        #print(c)
                        glis.append(c)

                    #glis.append('')
                    #glis.append('')
                    try:
                        u=driver.find_element_by_id('acrPopover').get_attribute('title')
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(u)
                        glis.append(u)



                    try:
                        #ama=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-0']/span[2]/span").get_attribute('innerHTML')
                        #ama2=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-1']/span[2]/span").get_attribute('innerHTML')
                        ama=driver.find_element_by_css_selector("#tabular-buybox-truncate-0 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        ama2=driver.find_element_by_css_selector("#tabular-buybox-truncate-1 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        soup1 = BeautifulSoup(ama,features="lxml").text
                        soup2 = BeautifulSoup(ama2,features="lxml").text
                        
                        #print(bdy)
                        glis.append('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        print('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        
                        
                    except Exception:
                        glis.append("not available")
                    else:
                        pass


                    try:
                        el=driver.find_element_by_id("availability").text
                        try:
                            for x in range(1,10):
                                el=el.replace("\n", " ")
                        except Exception:
                            pass
                    except Exception:
                        print(y)
                        glis.append(y)
                        el=y
                    else:
                        #print(el)
                        if el:
                            
                            glis.append(el)
                        else:
                            glis.append(y)
                            el=y


                    got=False
                    got2=False
                    fina='none'
                    dim=False
                    rep=re.split('\s+', finalhtml)
                    for t in rep:
                        if 'ounces' in t.lower():
                            try:
                                e=rep.index('ounces')
                                print(rep[e-1] +' '+'ounces')
                                glis.append(0.5)
                                dim=True
                                break
                            except Exception as r:
                                print(r)
                        else:
                            try:
                                if 'pounds' in t.lower():
                                    print(t)
                                    e=rep.index('pounds')
                                    print(rep[e-1]+' '+'pounds')
                                    glis.append(rep[e-1])
                                    dim=True
                                    break
                                else:
                                    pass
                            except Exception as p:
                                pass
                    if dim==True:
                        pass
                    else:
                        glis.append('not available')
                    inch=False  
                    date_avail=False  
                    for ter in rep:
                        if 'inches' in ter:
                            rep2=rep.index(ter)
                            if rep[rep2-2]=="x":
                                pu=rep[rep2-5]+rep[rep2-4]+rep[rep2-3]+rep[rep2-2]+rep[rep2-1]+' '+rep[rep2]
                                glis.append(pu)
                                inch=True
                                break
                    for ter in rep:
                        if inch==True:
                            break
                        if 'Dimensions' in ter:
                            rep2=rep.index(ter)
                            print(rep[rep2-2])
                            if rep[rep2+2]=="x":
                                pu=rep[rep2+1]+rep[rep2+2]+rep[rep2+3]+rep[rep2+4]+rep[rep2+5]
                                glis.append(pu)
                                inch=True
                                break
                    if inch==False:
                        glis.append(y)
                    for tor in rep:    
                        if 'Date' in tor:

                            rep2=rep.index(tor)
                            if "First" in rep[rep2+1]:
                                print(rep[rep2-1])
                                pu2=rep[rep2+3]+' '+rep[rep2+4]+' '+rep[rep2+5]
                                if ':' in pu2:
                                    break
                                glis.append(pu2)
                                date_avail=True
                                break

                    if ':' in pu2:
                        for tor in rep:    
                            if 'Date' in tor:

                                rep2=rep.index(tor)
                                if "First" in rep[rep2+1]:
                                    print(rep[rep2-1])
                                    pu2=rep[rep2+6]+' '+rep[rep2+7]+' '+rep[rep2+8]
                                    glis.append(pu2)
                                    date_avail=True
                                    break
            
                    if date_avail==False:
                        glis.append(y)
                    dr=[]
                    put=False
                    put2=False
                    for jt in range(1,3):
                        
                        try:
                            driver.execute_script("arguments[0].setAttribute('class', 'a-section a-spacing-none a-spacing-top-small a-padding-none')",driver.find_element_by_xpath('//*[@id="aod-pinned-offer-additional-content"]'))    
                            #driver.find_element_by_id("aod-pinned-offer-show-more-link").click()
                            time.sleep(2)
                            dr=driver.find_element_by_id('aod-pinned-offer').text.split('\n')
                            print(dr)
                            print('ok')
                            break
                        except Exception:
                            time.sleep(3)
                            print('ok')
                    try:
                        #print(driver.find_element_by_id("aod-pinned-offer-additional-content").text)
                    
                        print(dr)
                        try:
                            x=dr.index('New')
                        except Exception:
                            x=dr.index('New - Business Price')
                        print(dr[x+1]+'.'+dr[x+2])
                        glis.append(str(dr[x+1]+'.'+dr[x+2]))
                        yr=dr.index('Ships from')
                        print(' '.join(dr[yr:]))
                        glis.append(' '.join(dr[yr:]))
                        put2=True
                    except Exception:
                        pass


                    try:
                        driver.execute_script("document.getElementById('aod-filter-list').style.display = 'block';")
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("primeEligible"))
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("new"))
                        print(driver.find_element_by_id('aod-filter-list').text)
                        time.sleep(2)
                        fo=driver.find_elements_by_id("aod-offer")
                        iv=[]
                        for io in fo:
                            ok=io.text
                            iv.append(ok.split('\n'))
                        yo=0
                        print(iv)
                        for a in iv:
                        
                            x=iv[yo].index('New')
                            print(iv[yo][x+1]+'.'+iv[yo][x+2])
                            glis.append(str(iv[yo][x+1]+'.'+iv[yo][x+2]))
                            yr=iv[yo].index('Ships from')
                            print(' '.join(iv[yo][yr:]))
                            glis.append(' '.join(iv[yo][yr:]))
                            yo=yo+1
                            put=True
                    except Exception:
                        glis.append('NoNe')
                        glis.append('NoNEE')

                    m=1

                    glance=[]
                    
                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,90):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as klo:
                        print(klo)
                    try:
                        if put2==True:
                            if 'stock soon' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'months'in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'weeks' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'out of stock' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif '1 left' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'unavailable' in el.lower():
                                #signn['ASIN'][j]=col
                                sht1['H'+str(j)].value='NO'
                            elif 'released' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'preorder' in  el.lower():
                                sht1['H'+str(j)].value='NO'
                            else:
                                sht1['H'+str(j)].value='YES'
                        elif not iv:
                            sht1['H'+str(j)].value='NO'
                        else:
                            pass
                        if put==False:
                            if 'NoNe' in glis[10]:
                                sht1['H'+str(j)].value='NO'
                        else:
                            sht1['H'+str(j)].value='YES'
                    except Exception:
                        pass
                    if 'unavailable' in el.lower():
                        signn['ASIN'][j]=col

                    #nil=nil+1
                    #driver.quit
                    print(glis[3])
                    glis=[]
                    glance=[]
                else:
                    glance=[]
                    m=1
                    urr="https://www.amazon.com/dp/"+str(col)
                    #print(col)
                    glis=[]
                    
                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                    unavail['ASIN'].append(col)
                    unavail['SKU'].append(sht1['B'+str(j)].value)
                    #sheet.insert_row(glis,2)
                    

                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,67+len(xy)):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as kp:
                        print(kp)
                    sht1['H'+str(j)].value='NO'
                    glance=[]
                    
                    glis=[]
                #nil=nil+1
                try:
                    driver.find_element_by_xpath("//*[@id='nav-logo']/a/span")
                except Exception as ou:
                    print(ou)
                    
                    try:
                        driver.find_element_by_xpath("//*[@id='c']")
                    except Exception as og:
                        print(og)
                        if col=='None':
                            pass
                        elif empty==True:
                            pass
                        else:
                            print(ogg)
                    else:
                        try:
                            for l in range(67,67+len(xy)):

                                sht1[chr(l)+str(j)].value='unavailable'
                                m=m+1

                            print()
                        except Exception as kp:
                            print(kp)
                    
                else:
                    pass
                star=star+1

                
                #glance_called=False


                



            print()


        except Exception as hee:
            print('error here')
            star=safe
            try:
                
                driver.quit()
            except Exception:
                pass
            print(hee)
            print("xyxyxy")
            now = datetime.now()
    
            current_time = now.strftime("%I:%M:%p")
    
            
            pop13=["simple det error@ thread1",'index '+str(j),current_time]
            try:
                sheet_log = gc.open("All information").worksheet(gk)
                sheet_log.insert_row(pop13,4)
                #print('look')
            except Exception:
                pass
        else:
            
            
            try:
                
                driver.quit()
                
            except Exception:
                print('interrupted')
            else:
                print('completed')

                break
        finally:
            atlist.append(1)
            if len(atlist)==4:
                print('signn')
                print(signn)
                sign6()
                print('unavail')
                print(unavail)
                df=pd.DataFrame.from_dict(unavail,orient='index').transpose()
                print(df)
                if unavail['ASIN']:
                    #self.spapii()
                    pass
                else:
                    pass
            print("It's done finally")
             
def asin3():
    global openn
    print("got here2")
    global files2
    global star2
    global en2
    files=files2
    global lk
    star=int(star2)
    global lo
    en=int(en2)
    global assure
    #print(files)
    global gk
    #star=star
    global assure
    assure=False
    j=star
    options = webdriver.ChromeOptions() 

    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data4")
    #options.add_argument('--headless')
    if openn==True:
        pass
    else:
        options.add_argument('--headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    #driver= webdriver.Chrome(executable_path=r"driver2/chromedriver.exe",options=options)
    #driver.get("https:www.amazon.com/dp/B01N9QKU2S")
    #time.sleep(3)
    #print(driver.find_element_by_xpath("//*[@id='priceBadging_feature_div']/i").get_attribute("aria-label"))
    #driver.quit()
    y="unavailable"
    x="unavailable"
    xy=[1,1,1,1,1,1,1,1,1,1,1]
            
    wb=xw.Book("C:\\Users\\Public\\"+str(files))
    sht1=wb.sheets['New File']
    #assure=False
    #star=2
    safe=star
    global signn
    global getval
    getval=getval
    getval=False
    hal=1
    for iml in range(1,100):

        if getval==True:

            print("getvalll")
            
        
            break
        
        #options = webdriver.ChromeOptions()
        try:
            #options= webdriver.ChromeOptions()

            try:
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            except Exception:
                time.sleep(5)
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            else:
                pass
            
            for j in range(star+2,en+1,4):
                if hal%500==0:
                    time.sleep(300)
                hal=hal+1
                glis=[]
                glance=[]
                safe=j-2
                drive_call=False
                ydrive_call=False
                Zdrive_call=False
                try:
                    col=sht1['A'+str(j)].value
                except Exception:
                    print("XL")
                    #time.sleep(10)
                    print(expected)
                else:
                    pass
                print("driver3")
                print(col)
                try:

                    
                    if assure==True:
                        break
                    else:
                        for ite in range(1,30):
                            try:
                                driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                time.sleep(random.randint(3,7))
                                #time.sleep(2000)
                                
                                drive_call=True
                                time.sleep(1)
                            except (WebDriverException,TimeoutException) as w:
                                print(w)
                                time.sleep(3)
                                driver.refresh()
                            else:
                                break
                        print()

                        if drive_call==True:
                            pass
                        else:
                            
                            for ite in range(1,90):
                                time.sleep(10)
                                try:
                                    
                                    driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                    time.sleep(random.randint(3,7))
                                    #drive_call=True
                                    time.sleep(1)
                                except (WebDriverException,TimeoutException) as w:
                                    time.sleep(3)
                                    driver.refresh()
                                else:
                                    break
                            print()
                        
                    get_title = driver.title
                    asi_call=False
                    for prod in range(1,5):

                        try:
                            a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                            
                        except Exception as pro:
                            if not driver.title:
                                empty=True
                                break
                            elif "Page Not Found" in get_title:
                                break

                            else:
                                #print(y)
                                print(pro)
                                time.sleep(5)
                                a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                        else:
                            #print(a)
                            ASIN=False
                            try:
                                html= str(driver.page_source)
                                soup = BeautifulSoup(html,features="lxml")
                                finalhtml=soup.text.replace('\n', ' ')
                            except Exception as pro:
                                pass
                            else:
                                #print(a)
                                if str(col) in finalhtml:
                                    print(hammer)
                                elif 'ASIN' in finalhtml:
                                    pass
                                else:
                                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                                    try:
                                        xi=driver.find_elements_by_tag_name('tbody')
                                        
                                    except Exception:
                                        pass
                                    else:
                                        for jop in xi:
                                            
                                            if 'ASIN' in jop.get_attribute('innerHTML'):
                                                ASIN=True
                                                #print(jop.get_attribute('innerHTML'))
                                                if str(col) in jop.get_attribute('innerHTML'):
                                                    print(hammer2)
                                                    
                                                else:
                                                    pass

                                        if ASIN==True:
                                            pass
                                        else:
                                            print(hammer3)
                    
                except Exception as km:
                    #print('km')
                    print(km)

                    #action = ActionChains(driver)
                    #firstLevelMenu = driver.find_element_by_id("acrPopover")
                    #action.move_to_element(firstLevelMenu).perform()
                    #time.sleep(5)
                    
                        
                    u="https://www.amazon.com/dp/"+str(col)
                    #u=' '
                    #uv="=HYPERLINK("+"https://www.amazon.com/dp/"+str(col)+, "product link")
                    glis.append(u)
                    #print(glis)
                    try:
                        a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(a)
                        glis.append(a)



                    try:
                        b=driver.find_element_by_id("bylineInfo_feature_div").text
                    except Exception:
                        print(y)
                        glis.append(y)
                    else:
                        #print(b)
                        glis.append(b)
                    print('ok')



                    try:
                        c=driver.find_element_by_xpath("//*[@id='priceblock_ourprice']").text

                    except Exception:
                        try:
                            z=driver.find_element_by_xpath("//*[@id='priceblock_businessprice']").text
                            
                        except Exception:
                            try:
                                q=driver.find_element_by_xpath("//*[@id='priceblock_saleprice']").text
                            except Exception:
                                try:
                                    p=driver.find_element_by_xpath("//*[@id='priceblock_dealprice']").text 
                                except Exception:
                                    try:
                                        zp=driver.find_element_by_xpath("//*[@id='priceblock_saleprice_lbl']").text
                                    except Exception:
                                        #print("none")
                                        glis.append("none")
                                    else:
                                        #print(p)
                                        glis.append(zp)
                                else:
                                    glis.append(p)
                            else:
                                #print(q)
                                glis.append(q)
                        else:
                            glis.append(z)
                    else:
                        #print(c)
                        glis.append(c)

                    #glis.append('')
                    #glis.append('')
                    try:
                        u=driver.find_element_by_id('acrPopover').get_attribute('title')
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(u)
                        glis.append(u)



                    try:
                        #ama=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-0']/span[2]/span").get_attribute('innerHTML')
                        #ama2=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-1']/span[2]/span").get_attribute('innerHTML')
                        ama=driver.find_element_by_css_selector("#tabular-buybox-truncate-0 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        ama2=driver.find_element_by_css_selector("#tabular-buybox-truncate-1 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        soup1 = BeautifulSoup(ama,features="lxml").text
                        soup2 = BeautifulSoup(ama2,features="lxml").text
                        
                        #print(bdy)
                        glis.append('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        print('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        
                        
                    except Exception:
                        glis.append("not available")
                    else:
                        pass


                    try:
                        el=driver.find_element_by_id("availability").text
                        try:
                            for x in range(1,10):
                                el=el.replace("\n", " ")
                        except Exception:
                            pass
                    except Exception:
                        print(y)
                        glis.append(y)
                        el=y
                    else:
                        #print(el)
                        if el:
                            
                            glis.append(el)
                        else:
                            glis.append(y)
                            el=y


                    got=False
                    got2=False
                    fina='none'
                    dim=False
                    rep=re.split('\s+', finalhtml)
                    for t in rep:
                        if 'ounces' in t.lower():
                            try:
                                e=rep.index('ounces')
                                print(rep[e-1] +' '+'ounces')
                                glis.append(0.5)
                                dim=True
                                break
                            except Exception as r:
                                print(r)
                        else:
                            try:
                                if 'pounds' in t.lower():
                                    print(t)
                                    e=rep.index('pounds')
                                    print(rep[e-1]+' '+'pounds')
                                    glis.append(rep[e-1])
                                    dim=True
                                    break
                                else:
                                    pass
                            except Exception as p:
                                pass
                    if dim==True:
                        pass
                    else:
                        glis.append('not available')
                    inch=False 
                    date_avail=False   

                    for ter in rep:
                        if 'inches' in ter:
                            rep2=rep.index(ter)
                            if rep[rep2-2]=="x":
                                pu=rep[rep2-5]+rep[rep2-4]+rep[rep2-3]+rep[rep2-2]+rep[rep2-1]+' '+rep[rep2]
                                glis.append(pu)
                                inch=True
                                break
                            
                    for ter in rep:
                        if inch==True:
                            break
                        if 'Dimensions' in ter:
                            rep2=rep.index(ter)
                            print(rep[rep2-2])
                            if rep[rep2+2]=="x":
                                pu=rep[rep2+1]+rep[rep2+2]+rep[rep2+3]+rep[rep2+4]+rep[rep2+5]
                                glis.append(pu)
                                inch=True
                                break
                    if inch==False:
                        glis.append(y)
                    for tor in rep:    
                        if 'Date' in tor:

                            rep2=rep.index(tor)
                            if "First" in rep[rep2+1]:
                                print(rep[rep2-1])
                                pu2=rep[rep2+3]+' '+rep[rep2+4]+' '+rep[rep2+5]
                                if ':' in pu2:
                                    break
                                glis.append(pu2)
                                date_avail=True
                                break

                    if ':' in pu2:
                        for tor in rep:    
                            if 'Date' in tor:

                                rep2=rep.index(tor)
                                if "First" in rep[rep2+1]:
                                    print(rep[rep2-1])
                                    pu2=rep[rep2+6]+' '+rep[rep2+7]+' '+rep[rep2+8]
                                    glis.append(pu2)
                                    date_avail=True
                                    break
    
                    if date_avail==False:
                        glis.append(y)

                    put=False
                    put2=False
                    dr=[]
                    for jt in range(1,3):
                        
                        try:
                            driver.execute_script("arguments[0].setAttribute('class', 'a-section a-spacing-none a-spacing-top-small a-padding-none')",driver.find_element_by_xpath('//*[@id="aod-pinned-offer-additional-content"]'))    
                            #driver.find_element_by_id("aod-pinned-offer-show-more-link").click()
                            time.sleep(2)
                            dr=driver.find_element_by_id('aod-pinned-offer').text.split('\n')
                            print(dr)
                            print('ok')
                            break
                        except Exception:
                            time.sleep(3)
                            print('ok')
                    try:
                        
                        #print(driver.find_element_by_id("aod-pinned-offer-additional-content").text)
                        
                        print(dr)
                        try:
                            x=dr.index('New')
                        except Exception:
                            x=dr.index('New - Business Price')
                        print(dr[x+1]+'.'+dr[x+2])
                        glis.append(str(dr[x+1]+'.'+dr[x+2]))
                        yr=dr.index('Ships from')
                        print(' '.join(dr[yr:]))
                        glis.append(' '.join(dr[yr:]))
                        put2=True
                    except Exception:
                        pass


                    try:
                        driver.execute_script("document.getElementById('aod-filter-list').style.display = 'block';")
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("primeEligible"))
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("new"))
                        print(driver.find_element_by_id('aod-filter-list').text)
                        time.sleep(2)
                        fo=driver.find_elements_by_id("aod-offer")
                        iv=[]
                        for io in fo:
                            ok=io.text
                            iv.append(ok.split('\n'))
                        yo=0
                        print(iv)
                        for a in iv:
                        
                            x=iv[yo].index('New')
                            print(iv[yo][x+1]+'.'+iv[yo][x+2])
                            glis.append(str(iv[yo][x+1]+'.'+iv[yo][x+2]))
                            yr=iv[yo].index('Ships from')
                            print(' '.join(iv[yo][yr:]))
                            glis.append(' '.join(iv[yo][yr:]))
                            yo=yo+1
                            put=True
                    except Exception:
                        glis.append('NoNe')
                        glis.append('NoNEE')

                    m=1

                    glance=[]
                    
                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,90):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as klo:
                        print(klo)
                    print('glis 10'+glis[10])
                    try:
                        if put2==True:
                            if 'stock soon' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'months'in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'weeks' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'out of stock' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif '1 left' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'unavailable' in el.lower():
                                #signn['ASIN'][j]=col
                                sht1['H'+str(j)].value='NO'
                            elif 'released' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'preorder' in  el.lower():
                                sht1['H'+str(j)].value='NO'
                            else:
                                sht1['H'+str(j)].value='YES'
                        elif not iv:
                            sht1['H'+str(j)].value='NO'
                        else:
                            pass
                        if put==False:
                            if 'NoNe' in glis[10]:
                                sht1['H'+str(j)].value='NO'
                        else:
                            sht1['H'+str(j)].value='YES'
                    except Exception:
                        pass
                    if 'unavailable' in el.lower():
                        signn['ASIN'][j]=col
                    #nil=nil+1
                    #driver.quit
                    print(glis[3])
                    glis=[]
                    glance=[]
                else:
                    glance=[]
                    m=1
                    urr="https://www.amazon.com/dp/"+str(col)
                    #print(col)
                    glis=[]
                    
                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                    unavail['ASIN'].append(col)
                    unavail['SKU'].append(sht1['B'+str(j)].value)
                    #sheet.insert_row(glis,2)
                    

                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,67+len(xy)):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as kp:
                        print(kp)
                    sht1['H'+str(j)].value='NO'
                    glance=[]
                    
                    glis=[]
                #nil=nil+1
                try:
                    driver.find_element_by_xpath("//*[@id='nav-logo']/a/span")
                except Exception as ou:
                    print(ou)
                    
                    try:
                        driver.find_element_by_xpath("//*[@id='c']")
                    except Exception as og:
                        print(og)
                        if col=='None':
                            pass
                        elif empty==True:
                            pass
                        else:
                            print(ogg)
                    else:
                        try:
                            for l in range(67,67+len(xy)):

                                sht1[chr(l)+str(j)].value='unavailable'
                                m=m+1

                            print()
                        except Exception as kp:
                            print(kp)
                    
                else:
                    pass
                star=star+1

                
                #glance_called=False


                



            print()


        except Exception as hee:
            print('error here')
            star=safe
            try:
                
                driver.quit()
            except Exception:
                pass
            print(hee)
            print("xyxyxy")
            now = datetime.now()
    
            current_time = now.strftime("%I:%M:%p")
    
            
            pop13=["simple det error@ thread1",'index '+str(j),current_time]
            try:
                sheet_log = gc.open("All information").worksheet(gk)
                sheet_log.insert_row(pop13,4)
                #print('look')
            except Exception:
                pass
        else:
            
            
            try:
                
                driver.quit()
                
            except Exception:
                print('interrupted')
            else:
                print('completed')

                break
        finally:
            atlist.append(1)
            if len(atlist)==4:
                print('signn')
                print(signn)
                sign6()
                print('unavail')
                print(unavail)
                df=pd.DataFrame.from_dict(unavail,orient='index').transpose()
                print(df)
                if unavail['ASIN']:
                    #self.spapii()
                    pass
                else:
                    pass
            print("It's done finally")



def asin4():
    global openn
    print("got here2")
    global files2
    global star2
    global en2
    files=files2
    global lk
    star=int(star2)
    global lo
    en=int(en2)
    global assure
    #print(files)
    global gk
    #star=star
    global assure
    assure=False
    j=star
    options = webdriver.ChromeOptions() 

    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data5")
    #options.add_argument('--headless')
    if openn==True:
        pass
    else:
        options.add_argument('--headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    #driver= webdriver.Chrome(executable_path=r"driver2/chromedriver.exe",options=options)
    #driver.get("https:www.amazon.com/dp/B01N9QKU2S")
    #time.sleep(3)
    #print(driver.find_element_by_xpath("//*[@id='priceBadging_feature_div']/i").get_attribute("aria-label"))
    #driver.quit()
    y="unavailable"
    x="unavailable"
    xy=[1,1,1,1,1,1,1,1,1,1,1]
            
    wb=xw.Book("C:\\Users\\Public\\"+str(files))
    sht1=wb.sheets['New File']
    #assure=False
    #star=2
    safe=star
    global signn
    global getval
    getval=getval
    getval=False
    hal=1
    for iml in range(1,100):

        if getval==True:

            print("getvalll")
            
        
            break
        
        #options = webdriver.ChromeOptions()
        try:
            #options= webdriver.ChromeOptions()

            try:
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            except Exception:
                time.sleep(5)
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            else:
                pass
            
            for j in range(star+3,en+1,4):
                if hal%500==0:
                    time.sleep(300)
                hal=hal+1
                glis=[]
                glance=[]
                safe=j-3
                drive_call=False
                ydrive_call=False
                Zdrive_call=False
                try:
                    col=sht1['A'+str(j)].value
                except Exception:
                    print("XL")
                    #time.sleep(10)
                    print(expected)
                else:
                    pass
                print("driver3")
                print(col)
                try:

                    
                    if assure==True:
                        break
                    else:
                        for ite in range(1,30):
                            try:
                                driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                time.sleep(random.randint(3,7))
                                #time.sleep(2000)
                                
                                drive_call=True
                                time.sleep(1)
                            except (WebDriverException,TimeoutException) as w:
                                print(w)
                                time.sleep(3)
                                driver.refresh()
                            else:
                                break
                        print()

                        if drive_call==True:
                            pass
                        else:
                            
                            for ite in range(1,90):
                                time.sleep(10)
                                try:
                                    
                                    driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                    time.sleep(random.randint(3,7))
                                    #drive_call=True
                                    time.sleep(1)
                                except (WebDriverException,TimeoutException) as w:
                                    time.sleep(3)
                                    driver.refresh()
                                else:
                                    break
                            print()
                        
                    get_title = driver.title
                    asi_call=False
                    for prod in range(1,5):

                        try:
                            a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                            
                        except Exception as pro:
                            if not driver.title:
                                empty=True
                                break
                            elif "Page Not Found" in get_title:
                                break

                            else:
                                #print(y)
                                print(pro)
                                time.sleep(5)
                                a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                        else:
                            #print(a)
                            ASIN=False
                            try:
                                html= str(driver.page_source)
                                soup = BeautifulSoup(html,features="lxml")
                                finalhtml=soup.text.replace('\n', ' ')
                            except Exception as pro:
                                pass
                            else:
                                #print(a)
                                if str(col) in finalhtml:
                                    print(hammer)
                                elif 'ASIN' in finalhtml:
                                    pass
                                else:
                                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                                    try:
                                        xi=driver.find_elements_by_tag_name('tbody')
                                        
                                    except Exception:
                                        pass
                                    else:
                                        for jop in xi:
                                            
                                            if 'ASIN' in jop.get_attribute('innerHTML'):
                                                ASIN=True
                                                #print(jop.get_attribute('innerHTML'))
                                                if str(col) in jop.get_attribute('innerHTML'):
                                                    print(hammer2)
                                                    
                                                else:
                                                    pass

                                        if ASIN==True:
                                            pass
                                        else:
                                            print(hammer3)
                    
                except Exception as km:
                    #print('km')
                    print(km)

                    #action = ActionChains(driver)
                    #firstLevelMenu = driver.find_element_by_id("acrPopover")
                    #action.move_to_element(firstLevelMenu).perform()
                    #time.sleep(5)
                    
                        
                    u="https://www.amazon.com/dp/"+str(col)
                    #u=' '
                    #uv="=HYPERLINK("+"https://www.amazon.com/dp/"+str(col)+, "product link")
                    glis.append(u)
                    #print(glis)
                    try:
                        a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(a)
                        glis.append(a)



                    try:
                        b=driver.find_element_by_id("bylineInfo_feature_div").text
                    except Exception:
                        print(y)
                        glis.append(y)
                    else:
                        #print(b)
                        glis.append(b)
                    print('ok')



                    try:
                        c=driver.find_element_by_xpath("//*[@id='priceblock_ourprice']").text

                    except Exception:
                        try:
                            z=driver.find_element_by_xpath("//*[@id='priceblock_businessprice']").text
                            
                        except Exception:
                            try:
                                q=driver.find_element_by_xpath("//*[@id='priceblock_saleprice']").text
                            except Exception:
                                try:
                                    p=driver.find_element_by_xpath("//*[@id='priceblock_dealprice']").text 
                                except Exception:
                                    try:
                                        zp=driver.find_element_by_xpath("//*[@id='priceblock_saleprice_lbl']").text
                                    except Exception:
                                        #print("none")
                                        glis.append("none")
                                    else:
                                        #print(p)
                                        glis.append(zp)
                                else:
                                    glis.append(p)
                            else:
                                #print(q)
                                glis.append(q)
                        else:
                            glis.append(z)
                    else:
                        #print(c)
                        glis.append(c)

                    #glis.append('')
                    #glis.append('')
                    try:
                        u=driver.find_element_by_id('acrPopover').get_attribute('title')
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(u)
                        glis.append(u)



                    try:
                        #ama=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-0']/span[2]/span").get_attribute('innerHTML')
                        #ama2=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-1']/span[2]/span").get_attribute('innerHTML')
                        ama=driver.find_element_by_css_selector("#tabular-buybox-truncate-0 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        ama2=driver.find_element_by_css_selector("#tabular-buybox-truncate-1 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        soup1 = BeautifulSoup(ama,features="lxml").text
                        soup2 = BeautifulSoup(ama2,features="lxml").text
                        
                        #print(bdy)
                        glis.append('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        print('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        
                        
                    except Exception:
                        glis.append("not available")
                    else:
                        pass

                    try:
                        el=driver.find_element_by_id("availability").text
                        try:
                            for x in range(1,10):
                                el=el.replace("\n", " ")
                        except Exception:
                            pass
                    except Exception:
                        print(y)
                        glis.append(y)
                        el=y
                    else:
                        #print(el)
                        if el:
                            
                            glis.append(el)
                        else:

                            glis.append(y)
                            el=y


                    got=False
                    got2=False
                    fina='none'
                    dim=False
                    rep=re.split('\s+', finalhtml)
                    for t in rep:
                        if 'ounces' in t.lower():
                            try:
                                e=rep.index('ounces')
                                print(rep[e-1] +' '+'ounces')
                                glis.append(0.5)
                                dim=True
                                break
                            except Exception as r:
                                print(r)
                        else:
                            try:
                                if 'pounds' in t.lower():
                                    print(t)
                                    e=rep.index('pounds')
                                    print(rep[e-1]+' '+'pounds')
                                    glis.append(rep[e-1])
                                    dim=True
                                    break
                                else:
                                    pass
                            except Exception as p:
                                pass
                    if dim==True:
                        pass
                    else:
                        glis.append('not available')
                    inch=False    
                    date_avail=False
                    for ter in rep:
                        if 'inches' in ter:
                            rep2=rep.index(ter)
                            if rep[rep2-2]=="x":
                                pu=rep[rep2-5]+rep[rep2-4]+rep[rep2-3]+rep[rep2-2]+rep[rep2-1]+' '+rep[rep2]
                                glis.append(pu)
                                inch=True
                                break
                    for ter in rep:
                        if inch==True:
                            break
                        if 'Dimensions' in ter:
                            rep2=rep.index(ter)
                            print(rep[rep2-2])
                            if rep[rep2+2]=="x":
                                pu=rep[rep2+1]+rep[rep2+2]+rep[rep2+3]+rep[rep2+4]+rep[rep2+5]
                                glis.append(pu)
                                inch=True
                                break
                    if inch==False:
                        glis.append(y)
                    for tor in rep:    
                        if 'Date' in tor:

                            rep2=rep.index(tor)
                            if "First" in rep[rep2+1]:
                                print(rep[rep2-1])
                                pu2=rep[rep2+3]+' '+rep[rep2+4]+' '+rep[rep2+5]
                                if ':' in pu2:
                                    break
                                glis.append(pu2)
                                date_avail=True
                                break

                    if ':' in pu2:
                        for tor in rep:    
                            if 'Date' in tor:

                                rep2=rep.index(tor)
                                if "First" in rep[rep2+1]:
                                    print(rep[rep2-1])
                                    pu2=rep[rep2+6]+' '+rep[rep2+7]+' '+rep[rep2+8]
                                    glis.append(pu2)
                                    date_avail=True
                                    break
                            
                    
                    if date_avail==False:
                        glis.append(y)
                    
                    put=False
                    put2=False
                    dr=[]
                    for jt in range(1,3):
                        
                        try:
                            driver.execute_script("arguments[0].setAttribute('class', 'a-section a-spacing-none a-spacing-top-small a-padding-none')",driver.find_element_by_xpath('//*[@id="aod-pinned-offer-additional-content"]'))    
                            ##driver.find_element_by_id("aod-pinned-offer-show-more-link").click()
                            time.sleep(2)
                            dr=driver.find_element_by_id('aod-pinned-offer').text.split('\n')
                            print(dr)
                            print('ok')
                            break
                        except Exception:
                            time.sleep(3)
                            print('ok')
                    try:
                        #print(driver.find_element_by_id("aod-pinned-offer-additional-content").text)
                        
                        print(dr)
                        try:
                            x=dr.index('New')
                        except Exception:
                            x=dr.index('New - Business Price')
                        print(dr[x+1]+'.'+dr[x+2])
                        glis.append(str(dr[x+1]+'.'+dr[x+2]))
                        yr=dr.index('Ships from')
                        print(' '.join(dr[yr:]))
                        glis.append(' '.join(dr[yr:]))
                        put2=True
                    except Exception:
                        pass


                    try:
                        driver.execute_script("document.getElementById('aod-filter-list').style.display = 'block';")
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("primeEligible"))
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("new"))
                        print(driver.find_element_by_id('aod-filter-list').text)
                        time.sleep(2)
                        fo=driver.find_elements_by_id("aod-offer")
                        iv=[]
                        for io in fo:
                            ok=io.text
                            iv.append(ok.split('\n'))
                        yo=0
                        print(iv)
                        for a in iv:
                        
                            x=iv[yo].index('New')
                            print(iv[yo][x+1]+'.'+iv[yo][x+2])
                            glis.append(str(iv[yo][x+1]+'.'+iv[yo][x+2]))
                            yr=iv[yo].index('Ships from')
                            print(' '.join(iv[yo][yr:]))
                            glis.append(' '.join(iv[yo][yr:]))
                            yo=yo+1
                            put=True
                    except Exception:
                        glis.append('NoNe')
                        glis.append('NoNEE')

                    m=1

                    glance=[]
                    
                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,90):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as klo:
                        print(klo)
                    try:
                        if put2==True:
                            if 'stock soon' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'months'in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'weeks' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'out of stock' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif '1 left' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'unavailable' in el.lower():
                                #signn['ASIN'][j]=col
                                sht1['H'+str(j)].value='NO'
                            elif 'released' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'preorder' in  el.lower():
                                sht1['H'+str(j)].value='NO'
                            else:
                                sht1['H'+str(j)].value='YES'
                        elif not iv:
                            sht1['H'+str(j)].value='NO'
                        else:
                            pass
                        if put==False:
                            if 'NoNe' in glis[10]:
                                sht1['H'+str(j)].value='NO'
                        else:
                            sht1['H'+str(j)].value='YES'
                    except Exception:
                        pass
                    if 'unavailable' in el.lower():
                        signn['ASIN'][j]=col

                    #nil=nil+1
                    #driver.quit
                    print(glis[3])
                    glis=[]
                    glance=[]
                else:
                    glance=[]
                    m=1
                    urr="https://www.amazon.com/dp/"+str(col)
                    #print(col)
                    glis=[]
                    
                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                    unavail['ASIN'].append(col)
                    unavail['SKU'].append(sht1['B'+str(j)].value)
                    #sheet.insert_row(glis,2)
                    

                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,67+len(xy)):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as kp:
                        print(kp)
                    sht1['H'+str(j)].value='NO'
                    glance=[]
                    
                    glis=[]
                #nil=nil+1
                try:
                    driver.find_element_by_xpath("//*[@id='nav-logo']/a/span")
                except Exception as ou:
                    print(ou)
                    
                    try:
                        driver.find_element_by_xpath("//*[@id='c']")
                    except Exception as og:
                        print(og)
                        if col=='None':
                            pass
                        elif empty==True:
                            pass
                        else:
                            print(ogg)
                    else:
                        try:
                            for l in range(67,67+len(xy)):

                                sht1[chr(l)+str(j)].value='unavailable'
                                m=m+1

                            print()
                        except Exception as kp:
                            print(kp)
                    
                else:
                    pass
                star=star+1

                
                #glance_called=False


                



            print()


        except Exception as hee:
            print('error here')
            star=safe
            try:
                
                driver.quit()
            except Exception:
                pass
            print(hee)
            print("xyxyxy")
            now = datetime.now()
    
            current_time = now.strftime("%I:%M:%p")
    
            
            pop13=["simple det error@ thread1",'index '+str(j),current_time]
            try:
                sheet_log = gc.open("All information").worksheet(gk)
                sheet_log.insert_row(pop13,4)
                #print('look')
            except Exception:
                pass
        else:
            
            
            try:
                
                driver.quit()
                
            except Exception:
                print('interrupted')
            else:
                print('completed')

                break
        finally:
            atlist.append(1)
            if len(atlist)==4:
                print('signn')
                print(signn)
                sign6()
                print('unavail')
                print(unavail)
                df=pd.DataFrame.from_dict(unavail,orient='index').transpose()
                print(df)
                if unavail['ASIN']:
                    #self.spapii()
                    pass
                else:
                    pass
            print("It's done finally")

def sign6():
    global openn
    print("got here2")
    global files2
    global star2
    global en2
    files=files2
    global lk
    star=int(star2)
    global lo
    en=int(en2)
    global assure
    #print(files)
    global gk
    #star=star
    global assure
    assure=False
    j=star
    options = webdriver.ChromeOptions() 

    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data7")
    #options.add_argument('--headless')
    if openn==True:
        pass
    else:
        options.add_argument('--headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    #driver= webdriver.Chrome(executable_path=r"driver2/chromedriver.exe",options=options)
    #driver.get("https:www.amazon.com/dp/B01N9QKU2S")
    #time.sleep(3)
    #print(driver.find_element_by_xpath("//*[@id='priceBadging_feature_div']/i").get_attribute("aria-label"))
    #driver.quit()
    y="unavailable"
    x="unavailable"
    xy=[1,1,1,1,1,1,1,1,1,1,1]
            
    wb=xw.Book("C:\\Users\\Public\\"+str(files))
    sht1=wb.sheets['New File']
    #assure=False
    #star=2
    safe=star
    global signn
    global getval
    getval=getval
    getval=False
    hal=1
    ind=list(signn['ASIN'].keys())
    asinnt=list(signn['ASIN'].values())

    for iml in range(1,100):

        if getval==True:

            print("getvalll")
            
        
            break
        
        #options = webdriver.ChromeOptions()
        try:
            #options= webdriver.ChromeOptions()

            try:
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            except Exception:
                time.sleep(5)
                driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
                #driver.maximize_window()
                driver.set_page_load_timeout(60)
            else:
                pass
            
            for j in ind:
                if hal%500==0:
                    time.sleep(300)
                hal=hal+1
                glis=[]
                glance=[]
                safe=j
                drive_call=False
                ydrive_call=False
                Zdrive_call=False
                try:
                    col=sht1['A'+str(j)].value
                except Exception:
                    print("XL")
                    #time.sleep(10)
                    print(expected)
                else:
                    pass
                print("driver3")
                print(col)
                try:

                    
                    if assure==True:
                        break
                    else:
                        for ite in range(1,30):
                            try:
                                driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                time.sleep(random.randint(3,7))
                                #time.sleep(2000)
                                
                                drive_call=True
                                time.sleep(1)
                            except (WebDriverException,TimeoutException) as w:
                                print(w)
                                time.sleep(3)
                                driver.refresh()
                            else:
                                break
                        print()

                        if drive_call==True:
                            pass
                        else:
                            
                            for ite in range(1,90):
                                time.sleep(10)
                                try:
                                    
                                    driver.get("https://www.amazon.com/gp/offer-listing/"+str(col)+"/ref=dp_olp_NEW_mbc?ie=UTF8&condition=NEW")
                                    time.sleep(random.randint(3,7))
                                    #drive_call=True
                                    time.sleep(1)
                                except (WebDriverException,TimeoutException) as w:
                                    time.sleep(3)
                                    driver.refresh()
                                else:
                                    break
                            print()
                        
                    get_title = driver.title
                    asi_call=False
                    for prod in range(1,5):

                        try:
                            a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                            
                        except Exception as pro:
                            if not driver.title:
                                empty=True
                                break
                            elif "Page Not Found" in get_title:
                                break

                            else:
                                #print(y)
                                print(pro)
                                time.sleep(5)
                                a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                        else:
                            #print(a)
                            ASIN=False
                            try:
                                html= str(driver.page_source)
                                soup = BeautifulSoup(html,features="lxml")
                                finalhtml=soup.text.replace('\n', ' ')
                            except Exception as pro:
                                pass
                            else:
                                #print(a)
                                if str(col) in finalhtml:
                                    print(hammer)
                                elif 'ASIN' in finalhtml:
                                    pass
                                else:
                                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                                    try:
                                        xi=driver.find_elements_by_tag_name('tbody')
                                        
                                    except Exception:
                                        pass
                                    else:
                                        for jop in xi:
                                            
                                            if 'ASIN' in jop.get_attribute('innerHTML'):
                                                ASIN=True
                                                #print(jop.get_attribute('innerHTML'))
                                                if str(col) in jop.get_attribute('innerHTML'):
                                                    print(hammer2)
                                                    
                                                else:
                                                    pass

                                        if ASIN==True:
                                            pass
                                        else:
                                            print(hammer3)
                    
                except Exception as km:
                    #print('km')
                    print(km)

                    #action = ActionChains(driver)
                    #firstLevelMenu = driver.find_element_by_id("acrPopover")
                    #action.move_to_element(firstLevelMenu).perform()
                    #time.sleep(5)
                    
                        
                    u="https://www.amazon.com/dp/"+str(col)
                    #u=' '
                    #uv="=HYPERLINK("+"https://www.amazon.com/dp/"+str(col)+, "product link")
                    glis.append(u)
                    #print(glis)
                    try:
                        a=driver.find_element_by_xpath("//*[@id='productTitle']").text
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(a)
                        glis.append(a)



                    try:
                        b=driver.find_element_by_id("bylineInfo_feature_div").text
                    except Exception:
                        print(y)
                        glis.append(y)
                    else:
                        #print(b)
                        glis.append(b)
                    print('ok')



                    try:
                        c=driver.find_element_by_xpath("//*[@id='priceblock_ourprice']").text

                    except Exception:
                        try:
                            z=driver.find_element_by_xpath("//*[@id='priceblock_businessprice']").text
                            
                        except Exception:
                            try:
                                q=driver.find_element_by_xpath("//*[@id='priceblock_saleprice']").text
                            except Exception:
                                try:
                                    p=driver.find_element_by_xpath("//*[@id='priceblock_dealprice']").text 
                                except Exception:
                                    try:
                                        zp=driver.find_element_by_xpath("//*[@id='priceblock_saleprice_lbl']").text
                                    except Exception:
                                        #print("none")
                                        glis.append("none")
                                    else:
                                        #print(p)
                                        glis.append(zp)
                                else:
                                    glis.append(p)
                            else:
                                #print(q)
                                glis.append(q)
                        else:
                            glis.append(z)
                    else:
                        #print(c)
                        glis.append(c)

                    #glis.append('')
                    #glis.append('')
                    try:
                        u=driver.find_element_by_id('acrPopover').get_attribute('title')
                    except Exception:
                        #print(y)
                        glis.append(y)
                    else:
                        #print(u)
                        glis.append(u)



                    try:
                        #ama=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-0']/span[2]/span").get_attribute('innerHTML')
                        #ama2=driver.find_element_by_xpath("//*[@id='tabular-buybox-truncate-1']/span[2]/span").get_attribute('innerHTML')
                        ama=driver.find_element_by_css_selector("#tabular-buybox-truncate-0 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        ama2=driver.find_element_by_css_selector("#tabular-buybox-truncate-1 > span.a-truncate-cut > span").get_attribute('innerHTML')
                        soup1 = BeautifulSoup(ama,features="lxml").text
                        soup2 = BeautifulSoup(ama2,features="lxml").text
                        
                        #print(bdy)
                        glis.append('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        print('Ships from '+str(soup1)+' and '+'sold by '+str(soup2))
                        
                        
                    except Exception:
                        glis.append("not available")
                    else:
                        pass


                    try:
                        el=driver.find_element_by_id("availability").text
                        try:
                            for x in range(1,10):
                                el=el.replace("\n", " ")
                        except Exception:
                            pass
                    except Exception:
                        print(y)
                        glis.append(y)
                        el=y
                    else:
                        #print(el)
                        if el:
                            
                            glis.append(el)
                        else:
                            glis.append(y)
                            el=y


                    got=False
                    got2=False
                    fina='none'
                    dim=False
                    rep=re.split('\s+', finalhtml)
                    for t in rep:
                        if 'ounces' in t.lower():
                            try:
                                e=rep.index('ounces')
                                print(rep[e-1] +' '+'ounces')
                                glis.append(0.5)
                                dim=True
                                break
                            except Exception as r:
                                print(r)
                        else:
                            try:
                                if 'pounds' in t.lower():
                                    print(t)
                                    e=rep.index('pounds')
                                    print(rep[e-1]+' '+'pounds')
                                    glis.append(rep[e-1])
                                    dim=True
                                    break
                                else:
                                    pass
                            except Exception as p:
                                pass
                    if dim==True:
                        pass
                    else:
                        glis.append('not available')
                    inch=False    
                    date_avail=False
                    for ter in rep:
                        if 'inches' in ter:
                            rep2=rep.index(ter)
                            if rep[rep2-2]=="x":
                                pu=rep[rep2-5]+rep[rep2-4]+rep[rep2-3]+rep[rep2-2]+rep[rep2-1]+' '+rep[rep2]
                                glis.append(pu)
                                inch=True
                                break
                    for ter in rep:
                        if inch==True:
                            break
                        if 'Dimensions' in ter:
                            rep2=rep.index(ter)
                            print(rep[rep2-2])
                            if rep[rep2+2]=="x":
                                pu=rep[rep2+1]+rep[rep2+2]+rep[rep2+3]+rep[rep2+4]+rep[rep2+5]
                                glis.append(pu)
                                inch=True
                                break
                    if inch==False:
                        glis.append(y)
                    for tor in rep:    
                        if 'Date' in tor:
                            rep2=rep.index(tor)
                            if "First" in rep[rep2+1]:
                                print(rep[rep2-1])
                                pu2=rep[rep2+3]+' '+rep[rep2+4]+' '+rep[rep2+5]
                                if ':' in pu2:
                                    break
                                glis.append(pu2)
                                date_avail=True
                                break

                    if ':' in pu2:
                        for tor in rep:    

                            if 'Date' in tor:

                                rep2=rep.index(tor)
                                if "First" in rep[rep2+1]:
                                    print(rep[rep2-1])
                                    pu2=rep[rep2+6]+' '+rep[rep2+7]+' '+rep[rep2+8]
                                    glis.append(pu2)
                                    date_avail=True
                                    break
                    if date_avail==False:
                        glis.append(y)
                    
                    put=False
                    put2=False
                    dr=[]
                    for jt in range(1,3):
                        
                        try:
                            driver.execute_script("arguments[0].setAttribute('class', 'a-section a-spacing-none a-spacing-top-small a-padding-none')",driver.find_element_by_xpath('//*[@id="aod-pinned-offer-additional-content"]'))    
                            #driver.find_element_by_id("aod-pinned-offer-show-more-link").click()
                            time.sleep(2)
                            dr=driver.find_element_by_id('aod-pinned-offer').text.split('\n')
                            print(dr)
                            print('ok')
                            break
                        except Exception:
                            time.sleep(3)
                            print('ok')
                    try:
                        #print(driver.find_element_by_id("aod-pinned-offer-additional-content").text)
                    
                        print(dr)
                        try:
                            x=dr.index('New')
                        except Exception:
                            x=dr.index('New - Business Price')
                        print(dr[x+1]+'.'+dr[x+2])
                        glis.append(str(dr[x+1]+'.'+dr[x+2]))
                        yr=dr.index('Ships from')
                        print(' '.join(dr[yr:]))
                        glis.append(' '.join(dr[yr:]))
                        put2=True
                    except Exception:
                        pass


                    try:
                        driver.execute_script("document.getElementById('aod-filter-list').style.display = 'block';")
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("primeEligible"))
                        driver.execute_script("arguments[0].click();", driver.find_element_by_id("new"))
                        print(driver.find_element_by_id('aod-filter-list').text)
                        time.sleep(5)
                        fo=driver.find_elements_by_id("aod-offer")
                        iv=[]
                        for io in fo:
                            ok=io.text
                            iv.append(ok.split('\n'))
                        yo=0
                        print(iv)
                        for a in iv:
                            try:
                                x=iv[yo].index('New - Business Price')
                            except Exception:
                                x=iv[yo].index('New')
                            print(iv[yo][x+1]+'.'+iv[yo][x+2])
                            glis.append(str(iv[yo][x+1]+'.'+iv[yo][x+2]))
                            yr=iv[yo].index('Ships from')
                            print(' '.join(iv[yo][yr:]))
                            glis.append(' '.join(iv[yo][yr:]))
                            yo=yo+1
                            put=True
                    except Exception:
                        glis.append('NoNe')
                        glis.append('NoNEE')

                    m=1

                    glance=[]
                    
                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,90):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as klo:
                        print(klo)
                    print('el is '+el)
                    try:
                        if put2==True:
                            if 'stock soon' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'months'in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'weeks' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'out of stock' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif '1 left' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'unavailable' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'released' in el.lower():
                                sht1['H'+str(j)].value='NO'
                            elif 'preorder' in  el.lower():
                                sht1['H'+str(j)].value='NO'
                            else:
                                sht1['H'+str(j)].value='YES'
                        elif not iv:
                            sht1['H'+str(j)].value='NO'
                        else:
                            pass
                        if put==False:
                            if 'NoNe' in glis[10]:

                                sht1['H'+str(j)].value='NO'
                        else:
                            sht1['H'+str(j)].value='YES'
                    except Exception:
                        pass
                                
                        

                    #nil=nil+1
                    #driver.quit
                    print(glis[3])
                    glis=[]
                    glance=[]
                else:
                    glance=[]
                    m=1
                    urr="https://www.amazon.com/dp/"+str(col)
                    #print(col)
                    glis=[]
                    
                    glis.extend([urr,y,y,y,y,y,y,y,y,y,y,y,y,y,y])
                    unavail['ASIN'].append(col)
                    unavail['SKU'].append(sht1['B'+str(j)].value)
                    #sheet.insert_row(glis,2)
                    

                    #print("kkk is" + str(kkk))
                    try:
                        for l in range(67,67+len(xy)):

                            sht1[chr(l)+str(j)].value=glis[m-1]
                            m=m+1

                        print()
                    except Exception as kp:
                        print(kp)
                    sht1['H'+str(j)].value='NO'
                    glance=[]
                    
                    glis=[]
                #nil=nil+1
                try:
                    driver.find_element_by_xpath("//*[@id='nav-logo']/a/span")
                except Exception as ou:
                    print(ou)
                    
                    try:
                        driver.find_element_by_xpath("//*[@id='c']")
                    except Exception as og:
                        print(og)
                        if col=='None':
                            pass
                        elif empty==True:
                            pass
                        else:
                            print(ogg)
                    else:
                        try:
                            for l in range(67,67+len(xy)):

                                sht1[chr(l)+str(j)].value='unavailable'
                                m=m+1

                            print()
                        except Exception as kp:
                            print(kp)
                    
                else:
                    pass
                star=star+1

                
                #glance_called=False


                



            print()


        except Exception as hee:
            print('error here')
            star=safe
            try:
                
                driver.quit()
            except Exception:
                pass
            print(hee)
            print("xyxyxy")
            now = datetime.now()
    
            current_time = now.strftime("%I:%M:%p")
    
            
            pop13=["simple det error@ thread1",'index '+str(j),current_time]
            try:
                sheet_log = gc.open("All information").worksheet(gk)
                sheet_log.insert_row(pop13,4)
                #print('look')
            except Exception:
                pass
        else:
            
            
            try:
                
                driver.quit()
                
            except Exception:
                print('interrupted')
            else:
                print('completed')

                break
        finally:
            atlist.append(1)
            if len(atlist)==4:
                print('unavail')
                print(unavail)
                df=pd.DataFrame.from_dict(unavail,orient='index').transpose()
                print(df)
                if unavail['ASIN']:
                    #self.spapii()
                    pass
                else:
                    pass
            print("It's done finally")


def threader(files,star,en):
    global thr1
    global thr2
    global files2
    global star2
    global en2
    global thr3
    global thr4
    global signn
    global atlist
    atlist=[]
    signn={'ASIN':{}}
    files2=files
    star2=star
    en2=en
    os.system("TASKKILL /F /IM chromedriver.exe /T")
    thr1=threading.Thread(target=asin1)
    thr1.start()
    thr2=threading.Thread(target=asin2)
    thr2.start()
    thr3=threading.Thread(target=asin3)
    thr3.start()
    thr4=threading.Thread(target=asin4)
    thr4.start()

    

def chr1():
    print('came here')
    options = webdriver.ChromeOptions() 
    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data2")
    
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
    driver.get('https://www.amazon.com')
def chr2():
    options = webdriver.ChromeOptions() 
    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data3")
    
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
    driver.get('https://www.amazon.com')
def chr3():
    options = webdriver.ChromeOptions() 
    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data4")
    
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
    driver.get('https://www.amazon.com')
def chr4():
    options = webdriver.ChromeOptions() 
    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data5")
    
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
    driver.get('https://www.amazon.com')

def chr5():
    options = webdriver.ChromeOptions() 
    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data6")
    
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
    driver.get('https://www.amazon.com')

def chr6():
    options = webdriver.ChromeOptions() 
    options.add_argument("user-data-dir=C:\\Users\\Public\\simpledotcom\\new user data7")
    options.add_argument('window-size=1920x1080')
    options.add_argument('--no-proxy-server') 
    options.add_argument('log-level=3')
    options.add_experimental_option('w3c', False)
    driver= webdriver.Chrome(executable_path="driver2/chromedriver.exe", options=options)
    driver.get('https://www.amazon.com')

def hallt():  
    global getval
    getval=True
    global thr1
    global thr2
    global thr3
    global thr4
    
    thr1=thr1
    thr2=thr2
    thr3=thr3
    thr4=thr4
    
    os.system("TASKKILL /F /IM chromedriver.exe /T")

