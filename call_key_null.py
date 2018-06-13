#Call Key Messages with NULL Values Report
# Initial Coding : S.Mohana Priya
# Description: Automating the process of Updating Presentation Name and Presentation Id
#              for Call Keys having NULL Values in these fields.


import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import DataFrame, read_excel, merge, read_csv
import os
import openpyxl
import time
import win32com.client as win32

os.chdir('C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\Call_Key_Null_Msg')
timestr = time.strftime("%Y%m%d-%H%M%S")

# Extracting Data from Production
from simple_salesforce import Salesforce
import pandas as pd

sf = Salesforce(username='xxxx', password='yyyy',security_token='', sandbox=True)
query = sf.query_all("SELECT ID,PRESENTATION_ID_VOD__C,CLM_PRESENTATION_NAME_VOD__C,CLM_PRESENTATION_VOD__C from Call2_Key_Message_vod__c where CLM_PRESENTATION_NAME_VOD__C = NULL and CLM_PRESENTATION_VOD__C = NULL")
#query = sf.query_all("SELECT Id,Presentation_ID_vod__c,CLM_PRESENTATION_NAME_VOD__C,CLM_PRESENTATION_VOD__C from Call2_Key_Message_vod__c LIMIT 10")

ID = []
PRESENTATION_ID = []
#CLM_PRESENTATION_NAME = []


for record in query['records']:
    ID.append(record['Id'])
    PRESENTATION_ID.append(record['Presentation_ID_vod__c'] or '')

record_dict = {}
record_dict['ID'] = ID
record_dict['PRESENTATION_ID_VOD__C'] = PRESENTATION_ID

print(ID)  #List of Call_Key_Values
print(PRESENTATION_ID) #List of Presentation_Ids
print(record_dict) 

df = pd.DataFrame(record_dict)
df.to_excel('Call_Key_Null_Msg_ID'+timestr+'.xlsx',index=False)#----> contains Key ids and Presentation ids

n = len(PRESENTATION_ID)
print(n)


if (n == 0):
    print ("No Records with NULL Values")
else :
    print("Fetched %s records NULL Values" % n)

PRE_ID = []
PRE_NAME = []

for i in range(n):
    preId=PRESENTATION_ID[i]
    print ("Presentation Id is "+preId)
    query=sf.query_all("SELECT Id, Name from Clm_Presentation_vod__c where Id = '%s'" % preId)
    
    for record in query['records']:
        PRE_ID.append(record['Id'])
        PRE_NAME.append(record['Name'] or '')



final_dict = {}
final_dict['ID'] = ID
final_dict['CLM_PRESENTATION_NAME_VOD__C'] = PRE_NAME
final_dict['CLM_PRESENTATION_VOD__C'] = PRE_ID
final_dict['CLM_PRESENTATION_NAME_CELG__C'] = PRE_NAME
final_dict['OVERRIDE_LOCK_VOD__C'] = "TRUE"

df1 = pd.DataFrame(final_dict)
df2=df1[['ID','CLM_PRESENTATION_NAME_VOD__C','CLM_PRESENTATION_VOD__C','CLM_PRESENTATION_NAME_CELG__C','OVERRIDE_LOCK_VOD__C']]
df2.to_excel('Excel_To_Update'+timestr+'.xlsx',index=False)#---> Final Excel which is mailed to team. After verification can be uploaded via data loader. 


#df3 = pd.read_csv("C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\Call_Key_Null_Msg\\Excel_To_Update'+timestr+'.xlsx'

#Sending Update excel as mail attachment
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To='xyz.com'
mail.Subject='Update Excel for Call_Key Messages with NULL Values'
mail.Body='Hi Team ,\n Find Attched the Update Excel for Call Key Messages with Null Values.\n Kindly Verify and Update.\n Thanks.'
attachment = "C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\Call_Key_Null_Msg\\"+'Excel_To_Update'+timestr+'.xlsx'
mail.Attachments.Add(attachment)

mail.send
print("Update Excel Mailed to TEAM")

