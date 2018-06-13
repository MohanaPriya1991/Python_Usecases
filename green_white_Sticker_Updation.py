#Green Sticker/White Sticker Updation
#Initial Coding : S.Mohana Priya
#Description : Automating the process of Green/White Sticker Updation
#              Based on report from Business
# Updating Prodname,External id , User profiles.
#Input : Product Name.
#Altering Code - TO get Salesforce id as input

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

os.chdir('C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\Green_White_Update')
timestr = time.strftime("%Y%m%d-%H%M%S")

# Extracting Data from Production
from simple_salesforce import Salesforce
import pandas as pd

sf = Salesforce(username='data.admin@celgene.com.full2', password='celgene2',security_token='', sandbox=True)

#prod_name = str(sys.argv[1]) # Product Name.
#prod_name = input("Enter the Product Name : ")
sid =  input("Enter the Salesforce Id : ")  # Altering Code - TO get Salesforce id as input
print(sid)

query = sf.query_all("SELECT Id,Name,External_ID_vod__c from Product_vod__c where Id = '%s'" % sid)

ID = []
PROD_NAME = []
EXTERNAL_ID = []

for record in query ['records']:
    ID.append(record['Id'])
    PROD_NAME.append(record['Name'] or '')
    EXTERNAL_ID.append(record['External_ID_vod__c'] or '')

'''
id_len = len(ID)
for i in range (id_len):
    prod_id = ID[i]#-----> fetching Product Id
    print("The Product Id is : ",prod_id)
'''


def update_prod_name():
    print("Updating Product Name")
    new_prod_name = input("Enter the NEW Product Name : ")
    sf.Product_vod__c.update(sid,{'Name': new_prod_name})
    print("Product Name Updated")

def update_external_id():
    print("Updating External Id")
    new_ext_id = input("Enter the NEW External Id : ")
    sf.Product_vod__c.update(sid,{'External_ID_vod__c': new_ext_id,'VExternal_Id_vod__c':new_ext_id})
    print("External Id/Veeva External ID Updated")

def update_profile():
    print("Updating User Profile")
    old_prof = input("Enter the OLD Profile : ")
    new_prof = input("Enter the NEW Profile : ")
    #query = sf.query_all("SELECT Id from My_Setup_Products_vod__c where Product_vod__c = '%s'" % prod_id)
    query = sf.query_all("SELECT Id from My_Setup_Products_vod__c where Product_vod__c = '%s' and USER_PROFILE__C = '%s'" % (sid ,old_prof))
    
    MY_SET_UP_ID = [] #If no records exits
    

    for record in query ['records']:
        MY_SET_UP_ID.append(record['Id'])
        
    print(MY_SET_UP_ID)
    setup_id_len = len(MY_SET_UP_ID)
    print("No of set_up_ids : ",setup_id_len)

    if (setup_id_len == 0):
        print("No Records found in My_Setup_Products_vod__c for the given Product and User Profile")
    else :
        for j in range (setup_id_len):
            setId = MY_SET_UP_ID[j]
            sf.My_Setup_Products_vod__c.delete(setId)
            print("Deleted in My setup Products with ID : ",setId)

    query = sf.query_all("SELECT OwnerId from My_Setup_Products_vod__c where USER_PROFILE__C = '%s'" % new_prof)
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
        owner=unique_owner[i]
        pro = sf.My_Setup_Products_vod__c.create({'OwnerId':str(owner),
        'Product_vod__c':str(sid)})

    print("The Product is assigned with NEW USER_PROFILE : ",new_prof)



key = input("What is to be updated :\n A => Product Name\n B => External Id\n C => User Profile\n")
#print(key)

if (key == 'A'):
    update_prod_name() # Function call to update Product Name.
      
elif (key == 'B'):
    update_external_id() # Function call to update External Id.

elif (key == 'C'):
    update_profile() # Function call to update Profile.

else:
    print("Enter Valid Option")
    exit



    
