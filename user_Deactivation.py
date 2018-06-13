#User Deactivation
#Initial Coding : S.Mohana Priya
#Description : Automating the process of User Deactivation
#              as requested by Business.



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

os.chdir('C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\User_Deactivation')
timestr = time.strftime("%Y%m%d-%H%M%S")

# Extracting Data from Production
from simple_salesforce import Salesforce
import pandas as pd

sf = Salesforce(username='xxxx', password='yyyy',security_token='', sandbox=True)

username = input("Enter the USERNAME : ")
print(username)

query = sf.query_all("SELECT Id from User where Username = '%s'" % username)

USER_ID = []

for record in query ['records']:
    USER_ID.append(record['Id'])
       

id_len = len(USER_ID)
for i in range (id_len):
    USERID = USER_ID[i]#-----> fetching User Id
    print("The User Id to be Deleted is : ",USERID)


#update query to update IsActive to false and  removing UserRoleId
sf.User.update(USERID,{'UserRoleId': '','IsActive':'false'})

#print(query)
#print(sf.User.update(USERID,{'UserRole': ''}))

query = sf.query_all("SELECT Id,OwnerId from My_Setup_Products_vod__c where OwnerId = '%s'" % USERID)
SETUP_ID = []

for record in query ['records']:
    SETUP_ID.append(record['Id'])

setup_id_len = len(SETUP_ID) 
print("No of Records to be deleted from My_Setup_Products_vod__c : ",setup_id_len)
for i in range (setup_id_len):
    SETUPID = SETUP_ID[i]
    sf.My_Setup_Products_vod__c.delete(SETUPID) # deleteing Id's from My_Setup_Products_vod__c where ownerid matched the given userid.

print("USER DEACTIVATED")

# Public Group Membership
