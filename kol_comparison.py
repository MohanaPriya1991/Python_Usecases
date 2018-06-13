#KOL Comparison
#Initial Coding : S.Mohana Priya
#Description : Automating the process of statistical comparison of the
#              KOL data with the count received from Business
#Input  - Date --- Format --> 2018-01-04T00:00:00.000Z


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

os.chdir('C:\\Users\\m.priya.sankarrao\\Documents\\PYTHON\\KOL_Comparison')
timestr = time.strftime("%Y%m%d-%H%M%S")

# Extracting Data from Production
from simple_salesforce import Salesforce
import pandas as pd

sf = Salesforce(username='xxxx', password='yyyy',security_token='', sandbox=True)

#obj_name = input("Enter the Salesforce OBJECT Name : ")
kol_date = input("Enter the Date : ")

#print(obj_name)
#print(kol_date)

#query = sf.query_all("SELECT count() FROM Investigator_Information_CELG__c")

'''query = [SELECT count() FROM Investigator_Information_CELG__c];
print(query)
'''
'''
List<AggregateResult> result  = [select count() total from Account];
System.debug(result[0].get('total'));'''

#query = sf.query_all("SELECT Id FROM '%s' WHERE CreatedDate >= '%s'" % (obj_name ,kol_date))

print("###--------------------------------------OBJECT : Investigator_Information_CELG__c ---------------------------------###")
#---------No of Inserted Records---------------------#
query = sf.query_all("SELECT Id FROM Investigator_Information_CELG__c WHERE CreatedDate >= %s " % kol_date)

IN_ID_INSERT = [] 
    

for record in query ['records']:
    IN_ID_INSERT.append(record['Id'])
        
#print(IN_ID)
Inv_info_insert_count = len(IN_ID_INSERT)
print("No of Inserted Records in Investigator_Information_CELG__c : ",Inv_info_insert_count)


#---------No of Updated Records---------------------#
query = sf.query_all("SELECT Id FROM Investigator_Information_CELG__c WHERE LastModifiedDate >= %s " % kol_date)

IN_ID_UPDATE = [] 
    

for record in query ['records']:
    IN_ID_UPDATE.append(record['Id'])
        
#print(IN_ID)
Inv_info_update_count_tot = len(IN_ID_UPDATE) 
Inv_info_update_count = (Inv_info_update_count_tot - Inv_info_insert_count) # update count - insert count
print("No of Updated Records in Investigator_Information_CELG__c : ",Inv_info_update_count)


#------------No of Deleted Records----------------#
query = sf.query_all("SELECT Id FROM Investigator_Information_CELG__c WHERE IsDeleted = true AND LastModifiedDate >= %s " % kol_date)

IN_ID_DELETE = [] 
    

for record in query ['records']:
    IN_ID_DELETE.append(record['Id'])
        
#print(IN_ID)
Inv_info_delete_count = len(IN_ID_DELETE)
print("No of Deleted Records in Investigator_Information_CELG__c : ",Inv_info_delete_count)

print("###-----------------------------------------------------------------------------------------------------------------###")

print("###--------------------------------------OBJECT : Clinical_Trial__c ------------------------------------------------###")
#---------No of Inserted Records---------------------#
query = sf.query_all("SELECT Id FROM Clinical_Trial__c WHERE CreatedDate >= %s " % kol_date)

CLI_ID_INSERT = [] 
    

for record in query ['records']:
    CLI_ID_INSERT.append(record['Id'])
        
#print(IN_ID)
Cli_Trial_insert_count = len(CLI_ID_INSERT)
print("No of Inserted Records in Clinical_Trial__c : ",Cli_Trial_insert_count)


#---------No of Updated Records---------------------#
query = sf.query_all("SELECT Id FROM Clinical_Trial__c WHERE LastModifiedDate >= %s " % kol_date)

CLI_ID_UPDATE = [] 
    

for record in query ['records']:
    CLI_ID_UPDATE.append(record['Id'])
        
#print(IN_ID)
Cli_Trial_update_count_tot = len(CLI_ID_UPDATE) 
Cli_Trial_update_count = (Cli_Trial_update_count_tot - Cli_Trial_insert_count) # update count - insert count
print("No of Updated Records in Clinical_Trial__c : ",Cli_Trial_update_count)


#------------No of Deleted Records----------------#
query = sf.query_all("SELECT Id FROM Clinical_Trial__c WHERE IsDeleted = true AND LastModifiedDate >= %s " % kol_date)

CLI_ID_DELETE = [] 
    

for record in query ['records']:
    CLI_ID_DELETE.append(record['Id'])
        
#print(IN_ID)
Cli_Trial_delete_count = len(CLI_ID_DELETE)
print("No of Deleted Records in Clinical_Trial__c : ",Cli_Trial_delete_count)

print("###-----------------------------------------------------------------------------------------------------------------###")


print("###--------------------------------------OBJECT : Publication__c ---------------------------------------------------###")
#---------No of Inserted Records---------------------#
query = sf.query_all("SELECT Id FROM Publication__c WHERE CreatedDate >= %s " % kol_date)

PUB_ID_INSERT = [] 
    

for record in query ['records']:
    PUB_ID_INSERT.append(record['Id'])
        
#print(IN_ID)
Pub_insert_count = len(PUB_ID_INSERT)
print("No of Inserted Records in Publication__c : ",Pub_insert_count)


#---------No of Updated Records---------------------#
query = sf.query_all("SELECT Id FROM Publication__c WHERE LastModifiedDate >= %s " % kol_date)

PUB_ID_UPDATE = [] 
    

for record in query ['records']:
    PUB_ID_UPDATE.append(record['Id'])
        
#print(IN_ID)
Pub_update_count_tot = len(PUB_ID_UPDATE) 
Pub_update_count = (Pub_update_count_tot - Pub_insert_count) # update count - insert count
print("No of Updated Records in Publication__c : ",Pub_update_count)


#------------No of Deleted Records----------------#
query = sf.query_all("SELECT Id FROM Publication__c WHERE IsDeleted = true AND LastModifiedDate >= %s " % kol_date)

PUB_ID_DELETE = [] 
    

for record in query ['records']:
    PUB_ID_DELETE.append(record['Id'])
        
#print(IN_ID)
Pub_delete_count = len(PUB_ID_DELETE)
print("No of Deleted Records in Publication__c : ",Pub_delete_count)

print("###-----------------------------------------------------------------------------------------------------------------###")


print("###------------------------------------OBJECT : Publication_Involvement_CELG_House__c ------------------------------###")
#---------No of Inserted Records---------------------#
query = sf.query_all("SELECT Id FROM Publication_Involvement_CELG_House__c WHERE CreatedDate >= %s " % kol_date)

PUB_INV_ID_INSERT = [] 
    

for record in query ['records']:
    PUB_INV_ID_INSERT.append(record['Id'])
        
#print(IN_ID)
Pub_inv_insert_count = len(PUB_INV_ID_INSERT)
print("No of Inserted Records in Publication_Involvement_CELG_House__c : ",Pub_inv_insert_count)


#---------No of Updated Records---------------------#
query = sf.query_all("SELECT Id FROM Publication_Involvement_CELG_House__c WHERE LastModifiedDate >= %s " % kol_date)

PUB_INV_ID_UPDATE = [] 
    

for record in query ['records']:
    PUB_INV_ID_UPDATE.append(record['Id'])
        
#print(IN_ID)
Pub_inv_update_count_tot = len(PUB_INV_ID_UPDATE) 
Pub_inv_update_count = (Pub_inv_update_count_tot - Pub_inv_insert_count) # update count - insert count
print("No of Updated Records in Publication_Involvement_CELG_House__c : ",Pub_inv_update_count)


#------------No of Deleted Records----------------#
query = sf.query_all("SELECT Id FROM Publication_Involvement_CELG_House__c WHERE IsDeleted = true AND LastModifiedDate >= %s " % kol_date)

PUB_INV_ID_DELETE = [] 
    

for record in query ['records']:
    PUB_INV_ID_DELETE.append(record['Id'])
        
#print(IN_ID)
Pub_inv_delete_count = len(PUB_INV_ID_DELETE)
print("No of Deleted Records in Publication_Involvement_CELG_House__c : ",Pub_inv_delete_count)

print("###-----------------------------------------------------------------------------------------------------------------###")


print("###------------------------------------OBJECT : KOL_Profiler_Communities_CELG_House__c -----------------------------###")
#---------No of Inserted Records---------------------#
query = sf.query_all("SELECT Id FROM KOL_Profiler_Communities_CELG_House__c WHERE CreatedDate >= %s " % kol_date)

COM_ID_INSERT = [] 
    

for record in query ['records']:
    COM_ID_INSERT.append(record['Id'])
        
#print(IN_ID)
Com_insert_count = len(COM_ID_INSERT)
print("No of Inserted Records in KOL_Profiler_Communities_CELG_House__c : ",Com_insert_count)


#---------No of Updated Records---------------------#
query = sf.query_all("SELECT Id FROM KOL_Profiler_Communities_CELG_House__c WHERE LastModifiedDate >= %s " % kol_date)

COM_ID_UPDATE = [] 
    

for record in query ['records']:
    COM_ID_UPDATE.append(record['Id'])
        
#print(IN_ID)
Com_update_count_tot = len(COM_ID_UPDATE) 
Com_update_count = (Com_update_count_tot - Com_insert_count) # update count - insert count
print("No of Updated Records in KOL_Profiler_Communities_CELG_House__c : ",Com_update_count)


#------------No of Deleted Records----------------#
query = sf.query_all("SELECT Id FROM KOL_Profiler_Communities_CELG_House__c WHERE IsDeleted = true AND LastModifiedDate >= %s " % kol_date)

COM_ID_DELETE = [] 
    

for record in query ['records']:
    COM_ID_DELETE.append(record['Id'])
        
#print(IN_ID)
Com_delete_count = len(COM_ID_DELETE)
print("No of Deleted Records in KOL_Profiler_Communities_CELG_House__c : ",Com_delete_count)

print("###-----------------------------------------------------------------------------------------------------------------###")


print("###------------------------------------OBJECT : Community_Involvement_CELG_House__c --------------------------------###")
#---------No of Inserted Records---------------------#
query = sf.query_all("SELECT Id FROM Community_Involvement_CELG_House__c WHERE CreatedDate >= %s " % kol_date)

COM_INV_ID_INSERT = [] 
    

for record in query ['records']:
    COM_INV_ID_INSERT.append(record['Id'])
        
#print(IN_ID)
Com_inv_insert_count = len(COM_INV_ID_INSERT)
print("No of Inserted Records in Community_Involvement_CELG_House__c : ",Com_inv_insert_count)


#---------No of Updated Records---------------------#
query = sf.query_all("SELECT Id FROM Community_Involvement_CELG_House__c WHERE LastModifiedDate >= %s " % kol_date)

COM_INV_ID_UPDATE = [] 
    

for record in query ['records']:
    COM_INV_ID_UPDATE.append(record['Id'])
        
#print(IN_ID)
Com_inv_update_count_tot = len(COM_INV_ID_UPDATE) 
Com_inv_update_count = (Com_inv_update_count_tot - Com_inv_insert_count) # update count - insert count
print("No of Updated Records in Community_Involvement_CELG_House__c : ",Com_inv_update_count)


#------------No of Deleted Records----------------#
query = sf.query_all("SELECT Id FROM Community_Involvement_CELG_House__c WHERE IsDeleted = true AND LastModifiedDate >= %s " % kol_date)

COM_INV_ID_DELETE = [] 
    

for record in query ['records']:
    COM_INV_ID_DELETE.append(record['Id'])
        
#print(IN_ID)
Com_inv_delete_count = len(COM_INV_ID_DELETE)
print("No of Deleted Records in Community_Involvement_CELG_House__c : ",Com_inv_delete_count)

print("###-----------------------------------------------------------------------------------------------------------------###")


print("###------------------------------------OBJECT : Account ------------------------------------------------------------###")
#---------No of Inserted Records---------------------#
query = sf.query_all("SELECT Id FROM Account WHERE CreatedDate >= %s " % kol_date)

ACC_ID_INSERT = [] 
    

for record in query ['records']:
    ACC_ID_INSERT.append(record['Id'])
        
#print(IN_ID)
Acc_insert_count = len(ACC_ID_INSERT)
print("No of Inserted Records in Account : ",Acc_insert_count)


#---------No of Updated Records---------------------#
query = sf.query_all("SELECT Id FROM Account WHERE LastModifiedDate >= %s " % kol_date)

ACC_ID_UPDATE = [] 
    

for record in query ['records']:
    ACC_ID_UPDATE.append(record['Id'])
        
#print(IN_ID)
Acc_update_count_tot = len(ACC_ID_UPDATE) 
Acc_update_count = (Acc_update_count_tot - Acc_insert_count) # update count - insert count
print("No of Updated Records in Account : ",Acc_update_count)


#------------No of Deleted Records----------------#
query = sf.query_all("SELECT Id FROM Account WHERE IsDeleted = true AND LastModifiedDate >= %s " % kol_date)

ACC_ID_DELETE = [] 
    

for record in query ['records']:
    ACC_ID_DELETE.append(record['Id'])
        
#print(IN_ID)
Acc_delete_count = len(ACC_ID_DELETE)
print("No of Deleted Records in Account : ",Acc_delete_count)

print("###-----------------------------------------------------------------------------------------------------------------###")


