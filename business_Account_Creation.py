# Business Account Creation
# Initial Coding : S.Mohana Priya
# Description: Will automate the task of changing the RecordType of
#              an Account to "Business Account" if status
#              is "Approved by HomeOffice"

import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import DataFrame, read_excel, merge, read_csv
import os
import openpyxl
import time

os.chdir('C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\Business_Accout_Creation')
timestr = time.strftime("%Y%m%d-%H%M%S")

# Extracting Data from Production
from simple_salesforce import Salesforce
import pandas as pd

sf = Salesforce(username='xxxx', password='yyyy',security_token='', sandbox=True)

#query = sf.query_all("SELECT Id, Account_Status_CELG__c,recordtype.name from account where recordtype.name ='New Business Account' and Account_Status_CELG__c = 'Approved by Home Office' ")
query = sf.query_all("SELECT Id, Account_Status_CELG__c, RecordTypeId,Territory_vod__c from account where RecordTypeId ='012U00000001foq' and Account_Status_CELG__c = 'Approved by Home Office' ")


ID = []
STATUS = []
RECORD_TYPE = []
TERRITORY = []

#print(query)


for record in query['records']:
    ID.append(record['Id'])
    STATUS.append(record['Account_Status_CELG__c'] or '')
    RECORD_TYPE.append(record['RecordTypeId'] or '')
    TERRITORY.append(record['Territory_vod__c'] or '')
    #TERRITORIES.append(record[''] or '')

#print(ID)
record_dict = {}
record_dict['ID'] = ID
record_dict['ACCOUNT_STATUS_CELG__C'] = STATUS
record_dict['RecordTypeId'] = RECORD_TYPE
record_dict['Territory_vod__c'] = TERRITORY

df = pd.DataFrame(record_dict)
df.to_excel('Business_Account_BEFORE_'+timestr+'.xlsx')


n = len(ID)
#print(n)

rec_type = '012U00000001fop' # Id for Business Account

if (n == 0):
    print ("No Records with Status --APPROVED BY HOME OFFICE")
else :
    print("Fetched %s records with status--APPROVED BY HOME OFFICE" % n)

ID_AFTER = []
REC_TYPE_AFTER = []
TERRI = []

for i in range(n):
    accId=ID[i]
    print ("ID is "+accId)
    sf.Account.update(accId,{'RECORDTYPEID': rec_type}) #Updating from "New Business Account" to "Business Account" 
    #time.sleep(10)
    #query = sf.query_all("SELECT Id, RecordTypeId, Territory_vod__c from account where Id = '%s'" % accId)


for i in range(n):
    accId=ID[i]
    query = sf.query_all("SELECT Id, RecordTypeId, Territory_vod__c from account where Id = '%s'" % accId)

    for record in query['records']:
        ID_AFTER.append(record['Id'])
        REC_TYPE_AFTER.append(record['RecordTypeId'] or '')
        TERRI.append(record['Territory_vod__c'] or '')


record_dict1 = {}
record_dict1['ID'] = ID
record_dict1['RecordTypeId'] = REC_TYPE_AFTER
record_dict1['Territory_vod__c'] = TERRI

df1 = pd.DataFrame(record_dict1)
df1.to_excel('Business_Account_AFTER_'+timestr+'.xlsx')
    

#query = sf.query_all("SELECT Id, Account_Status_CELG__c, RecordTypeId,Territory_vod__c from account where Id in %s", % str(ID)
#query = sf.query_all("SELECT Id, Account_Status_CELG__c, RecordTypeId,Territory_vod__c from account where Id in (' + ','.join(map(str, Id)) + ')")
                     
                     
#query = sf.query_all("SELECT Id, Account_Status_CELG__c,recordtype.name from account where recordtype.name ='New Business Account' and Account_Status_CELG__c = 'Approved by Home Office' ")
#ID_AFTER = []

#for record in query['records']:
 #   ID_AFTER.append(record['Id'])
    #TERRITORIES.append(record[''] or '')

#print(ID_AFTER)'''
