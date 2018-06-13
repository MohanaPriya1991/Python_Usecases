#Green Sticker/White Sticker Creation
#Initial Coding : S.Mohana Priya
#Description : Automating the Process of Green/White Sticker Creation
#              Based on report from Business.

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

os.chdir('C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\Green_White_Create')
timestr = time.strftime("%Y%m%d-%H%M%S")

# Extracting Data from Production
from simple_salesforce import Salesforce
import pandas as pd

sf = Salesforce(username='xxxx', password='yyyy',security_token='', sandbox=True)

def create_green():
    PROD_NAME = input("Enter the Product Name : ")
    EXT_ID = input("Enter the External Id : ")
    TYPE = input("Enter the Product Type : ")

    prod = sf.Product_vod__c.create({'Name':str(PROD_NAME),
    'External_ID_vod__c':str(EXT_ID),
    'Product_Type_vod__c':str(TYPE),
    'Parent_Product_vod__c':'a00U000000RYKUcIAP',                                 
    'VExternal_Id_vod__c':str(EXT_ID)})
    print("New Product Created")
    #invoking  Apex Trigger-SampleLimitSet



def create_white():
    PROD_NAME = input("Enter the Product Name : ")
    EXT_ID = input("Enter the External Id : ")
    TYPE = input("Enter the Product Type : ")

    prod = sf.Product_vod__c.create({'Name':str(PROD_NAME),
    'External_ID_vod__c':str(EXT_ID),
    'Product_Type_vod__c':str(TYPE),
    'Parent_Product_vod__c':'a00U000000YQeydIAD',
    'VExternal_Id_vod__c':str(EXT_ID)})
    print("New Product Created")


def create_promotional():
    PROD_NAME = input("Enter the Product Name : ")
    EXT_ID = input("Enter the External Id : ")
    TYPE = input("Enter the Product Type : ")

    prod = sf.Product_vod__c.create({'Name':str(PROD_NAME),
    'External_ID_vod__c':str(EXT_ID),
    'Product_Type_vod__c':str(TYPE),
    'Parent_Product_vod__c':'a00U000000RYKUmIAP',
    'VExternal_Id_vod__c':str(EXT_ID)})
    print("New Product Created")


def create_user_profile():
    PROD_NAME = input("Enter the Product NAME : ") 
    NEW_PROF = input("Enter the USER Profile : ")

    query = sf.query_all("SELECT Id from Product_vod__c where Name = '%s'" % PROD_NAME)

    ID = []

    for record in query ['records']:
        ID.append(record['Id'])
       

    id_len = len(ID)
    for i in range (id_len):
        PROD_ID = ID[i]#-----> fetching Product Id
        print("The Product Id is : ",PROD_ID)
    
    query = sf.query_all("SELECT OwnerId from My_Setup_Products_vod__c where USER_PROFILE__C = '%s'" % NEW_PROF)
    OWNER_ID = []
    #print(query)
    for record in query ['records']:
        OWNER_ID.append(record['OwnerId'])

    owner_id_len = len(OWNER_ID)
    print("No of owner ids is : ",owner_id_len)

    #Removing Duplicat OwnerIds
    # insert the list to the set
    list_set = set(OWNER_ID)
    # convert the set to the list
    unique_owner = (list(list_set))

    uniq_len = len(unique_owner)
    print("No Of Unique Owner_ids is : ",uniq_len)


    #inserting Product_vod__c to distict ownerids of new user profile
    for i in range(uniq_len):
        OWNER=unique_owner[i]
        pro = sf.My_Setup_Products_vod__c.create({'OwnerId':str(OWNER),
        'Product_vod__c':str(PROD_ID)})

    print("The Product is assigned with USER_PROFILE : ",NEW_PROF)

    

key = input("Choose Whether GREEN / WHITE / PROMOTIONAL STICKER :\n A = > GREEN\n B => WHITE\n C => PROMOTIONAL\n D => USER PROFILE CREATION\n")
if (key =='A'):
    #print("create_green()")
    create_green() # Function call to Create GREEN STICKER Product.
    
elif (key == 'B'):
    #print("create_white()")
    create_white() # Function call to Create WHITE STICKER Product.

elif (key =='C'):
    create_promotional() # Function call to Create PROMOTIONAL STICKER Product.

elif (key == 'D'):
    create_user_profile()

else :
    print("Enter Valid Option")
    exit
