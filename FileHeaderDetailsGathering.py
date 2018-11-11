# -*- coding: utf-8 -*-
"""
Created on Sun Nov 11 19:51:29 2018

Find out important file header details of all the files in a given folder including its subfolders 

@author: Kelum Perera
"""


import pandas as pd
import os
import datetime as dt
import hashlib
import glob
from pandas import ExcelWriter
from pathlib import Path

# Input the desired folder to examine and the folder where the result can be saved

FolderPathToCheck = r'D:\Folder\ToCheck'
FolderPathToSaveResult = r'D:\SomeFolder\ToSaveResult'

# Delete previous objects, if any
del file_data
del file_data1

# Creates a pandas data frame        
file_Columns =['File_Path_Name','File_Size(Bytes)','Created_Time','Modified_Time','User_Info','Hash_md5']
file_data = pd.DataFrame(columns=file_Columns)

# Walkthrough the folder path and gather details and put into pandas data frame       

for root, dirs, files in os.walk(FolderPath):
    for fn in files:
        pathName = os.path.join(root, fn)            # Full path
        size = os.stat(pathName).st_size             # in bytes
        created_time = dt.datetime.fromtimestamp(os.stat(pathName).st_ctime)
        modified_time = dt.datetime.fromtimestamp(os.stat(pathName).st_mtime)
        userinfo = os.stat(pathName).st_uid          # in windows it always states as 0
        hash_md5 = hashlib.md5(open(os.path.join(root, fn), 'rb').read()).hexdigest()
        file_List =[[pathName,size,created_time,modified_time,userinfo,hash_md5]]
        file_data1 = pd.DataFrame(file_List,columns=file_Columns)
        file_data = file_data.append(file_data1,ignore_index=True)


# Split the "File_Path_Name" column into seperate columns as File Name, File Path and File extention        

for idx, row in file_data.iterrows():
    file_data.loc[idx,'File_Path'], file_data.loc[idx,'File_Name'] = os.path.split(file_data.loc[idx,'File_Path_Name'])
    file_data.loc[idx,'File_Extention'] =  Path(file_data.loc[idx,'File_Path_Name']).suffix
    
    

# Write the data frame to a excel sheet.

writer = ExcelWriter(FolderPathToSaveResult+"/All_File_Header_Data.xlsx")
file_data.to_excel(writer,'Sheet1',index=False)
writer.save()
