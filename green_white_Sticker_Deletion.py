#Green Sticker/White Sticker Deletion
#Initial Coding : S.Mohana Priya
#Description : Automating the process of Green/White Sticker deletion
#              Based on Green sticker Report from Business
#input : External ID.
# Altering Code - TO get Salesforce id as input


import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import DataFrame, read_excel, merge, read_csv
import os
import openpyxl
import time
import win32com.client as win32
import sys

os.chdir('C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\Green_White_Delete')
timestr = time.strftime("%Y%m%d-%H%M%S")

# Extracting Data from Production
from simple_salesforce import Salesforce
import pandas as pd

sf = Salesforce(username='data.admin@celgene.com.full2', password='celgene2',security_token='', sandbox=True)

#print ("This is the name of the script: ", sys.argv[0])
#print ("Number of arguments: ", len(sys.argv))
#print ("The arguments are: " , str(sys.argv[1]))

#prod_Name = str(sys.argv[1]) # first Argument - Product_name
#ext_id = str(sys.argv[1]) # Second Argument - External ID

#ext_id = input("Enter the External Id : ")
sid =  input("Enter the Salesforce Id : ")  # Altering Code - TO get Salesforce id as input

#print(prod_Name)
print(sid)
query = sf.query_all("SELECT Id,Name,External_ID_vod__c from Product_vod__c where Id = '%s'" % sid)

ID = []
PROD_NAME = []
EXTERNAL_ID = []

for record in query ['records']:
#    PRO_ID = record['Id'])
    ID.append(record['Id'])
    PROD_NAME.append(record['Name'] or '')
    EXTERNAL_ID.append(record['External_ID_vod__c'] or '')

name = len(PROD_NAME)
print("no is ",name)

#print("ID is ",ID)
'''id_len = len(ID)
for i in range (id_len):
    prod_id = ID[i]#-----> fetching Product Id
    print(prod_id)
  '''  


if (name == 0):
    print("No Product Name found for the given External ID")
    print("Please Contact Business Team for further information")
else :
    print("Please Verify the Details Below:")
    print("PRODUCT NAME : ",PROD_NAME)
    print("EXTERNAL ID : ",EXTERNAL_ID)
    print("Please Report to Business Team for Mismatch betweem PRODUCT NAME and EXTERNAL ID")

key = input("Do you want to continue -> Y/N : ")

if (key == 'N'):
    exit
elif (key =='Y'):
    query = sf.query_all("SELECT Id from My_Setup_Products_vod__c where Product_vod__c = '%s'" % sid)

    MY_SET_UP_ID = []

    for record in query ['records']:
        MY_SET_UP_ID.append(record['Id'])


    setup_id_len = len(MY_SET_UP_ID)
    print("No of set_up_ids : ",setup_id_len)
    #print(MY_SET_UP_ID)

    if (setup_id_len == 0):
        print("No Records found in My_Setup_Products_vod__c for the given Product Calatogue Id")
    else :
        for j in range (setup_id_len):
            setId = MY_SET_UP_ID[j]
            #print(setId)
            sf.My_Setup_Products_vod__c.delete(setId)

        print("Product Deleted")
    #   query=sf.query_all("DELETE from My_Setup_Products_vod__c where Id = '%s'" % setId)

else :
    print("Enter Valid Option")
    exit







