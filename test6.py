from pandas import ExcelWriter
import json
import string    
import random 
import pandas as pd
from pandas import *
import os,sys
import time
from openpyxl import Workbook
wb = Workbook()
import xlwings as xw
#st=str(input())
if os.path.isfile('C:\\Users\\Public\\Profit Price calculate.xlsx'):
    wb1=xw.Book("C:\\Users\\Public\\Profit Price calculate.xlsx")
    sht1=wb1.sheets['NEW FILE']
    sht1['L1'].value = "TRUE"


else:
    
    ws =  wb.active
    ws.title = "NEW FILE"
    wb.save(filename = 'C:\\Users\\Public\\Profit Price calculate.xlsx')
    wb1=xw.Book("C:\\Users\\Public\\Profit Price calculate.xlsx")
    sht1=wb1.sheets['NEW FILE']
    sht1['L1'].value = "TRUE"


print("welcome to profit price calculator session")
print("Before you run the program, please insert the data in required fields of the shown excel file (Profit price calculate.xlsx).\n Make sure the cells are not left empty.")
#print("\n When done please press 'Enter' to genarate random key")
S = 4
ra=random.choices(string.ascii_uppercase + string.digits, k = S)
ran = ','.join(ra)   
ran1 = ''.join(ra)    

print("Enter the string you see without commas(,) : " + str(ran)) # print the random data  
ni=str(input())
if ni==ran1:
    print('success')
else:
    print('fail')
    print(ran1)
    print('please rerun the program, Thankyou')
    time.sleep(5)
    sys.exit()
print('press "enter" again to confirm to run the task.')
inpp=input()
print('please wait the program is initiated this make take a few seconds, you can view the results in "NEW PRICE" column of the file.')
sht1['A1'].value ='ASINS'	
sht1['B1'].value ='seller-sku'	
sht1['C1'].value ='NEW PRICE'
sht1['D1'].value ='OLD PRICE'	
sht1['E1'].value ='Amazon Fee'	
sht1['F1'].value ='Amazon Fee %'
sht1['H1'].value ='Amazon/BrandOwner price'	
sht1['I1'].value ='weight in kg'	
sht1['J1'].value ='PROFIT'
sht1['K1'].value ='PROFIT %'
wb1.save('C:\\Users\\Public\\Profit Price calculate.xlsx')

xls = ExcelFile(r"C:\Users\Public\Profit Price calculate.xlsx")
df = xls.parse(xls.sheet_names[0])
data=df.to_dict()
#print(data)

with open('act.json', 'w') as outfile:
    json.dump(data, outfile)

df = pd.read_json('act.json')

print(df)
df1=len(df)

#print(df["NEW PRICE"][0])
for nu in range(0,df1):
    try:
        
        Jr=float(df["PROFIT"][nu])
        Dr=float(df["OLD PRICE"][nu])
        
        Fr=float(df["Amazon Fee %"][nu])
        Hr=float(df["Amazon/BrandOwner price"][nu])
        Ir=float(df["weight in kg"][nu])
        Kr=float(df["PROFIT %"][nu])
        
        for i in range(1,10000):
        
            Jr=Dr/1.18-Fr/100*Dr-7/100*Dr-200-200*Ir-Hr*1.2*75-Ir*4*75
            #print(J2)
            #print(J2)
            kr=Jr/(Dr-Jr)*100
            #print(k2)
            if kr<10:
                Dr=Dr+100
            elif kr>15:
                Dr=Dr-100
            else:
                #print('dr  is ')
                #print(Dr)
                break
        for p in range(1,10000):  
            if Jr<700:
                Dr=Dr+100
                Jr=Dr/1.18-Fr/100*Dr-7/100*Dr-200-200*Ir-Hr*1.2*75-Ir*4*75
            
                
            else:
                break

        data["NEW PRICE"][nu]=str(Dr)
        #print(data["NEW PRICE"][nu])
        #df["NEW PRICE"][nu]=str(Dr)
        df.at[nu,"NEW PRICE"]= Dr
        #df.loc["NEW PRICE"][nu] = Dr
        #break
        
        sht1['C'+str(nu+2)].value =Dr
    except Exception as ok:
        print(ok)
        break                            
#print('task completed, Thankyou')
#time.sleep(30)
#df.to_excel(r"C:\Users\Public\resultt.xlsx")
#time.sleep(2)



print("The run is successful")
print("if you want to rerun please close and open the program again ")


def func():
    print('its a function you know')