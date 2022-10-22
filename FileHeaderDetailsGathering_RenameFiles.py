# -*- coding: utf-8 -*-
"""
Created on Tue Nov 13 10:20:51 2018

Getting a data from excelsheet with a list new file names and old file names. ( or Create new file names)
Find out important file header details of all the files in a given folder including its subfolders save to an excelwork sheet
Getting a data from excelsheet with a list new file names and old file names 
Using the data from excel sheet, rename all the file names in a particular folder

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

FolderPathToCheck = r'D:\YourFolderPathHavingFiles\Test'
FolderPathToSaveResult = r'D:\YourFolderPathToSaveFiles\Test'


# Delete previous objects, if any
del file_data
del file_data1

excel_data = pd.read_excel("D:/YourFolderPathForExcelFiles/ExcelFile.xlsx", sheetname='Sheet1')

for idx, row in excel_data.iterrows():
    excel_data.loc[idx,'FolderPath'], excel_data.loc[idx,'FileName'] = os.path.split(excel_data.loc[idx,'Source_File_Path'])
    excel_data.loc[idx,'FileExtention'] =  Path(excel_data.loc[idx,'Source_File_Path']).suffix
    #df.loc[idx,'FileRenamed'] = df.loc[idx,'SN'] + df.loc[idx,'FileName']

excel_data['FileRenamed'] = excel_data['SN'].map(str) +" "+ excel_data['FileName']


# Creates a pandas data frame        
file_Columns =['File_Path_Name','File_Path','File_Name','File_Extention','File_Size(Bytes)','Created_Time','Modified_Time','User_Info','Hash_md5']
filesInFolder_data = pd.DataFrame(columns=file_Columns)

# Walkthrough the folder path and gather details and put into pandas data frame       

for root, dirs, files in os.walk(FolderPathToCheck):
    for fn in files:
        pathName = os.path.join(root, fn)            # Full path
        File_Path, File_Name = os.path.split(pathName) # Split the "Path_Name" column into seperate columns as File Name, File Path
        File_Extention =  Path(pathName).suffix       # Get the file extention from the "Path_Name" column
        size = os.stat(pathName).st_size             # in bytes
        created_time = dt.datetime.fromtimestamp(os.stat(pathName).st_ctime) # Gate created date time
        modified_time = dt.datetime.fromtimestamp(os.stat(pathName).st_mtime) # Gate modified date time
        userinfo = os.stat(pathName).st_uid          # in windows it always states as 0
        hash_md5 = hashlib.md5(open(os.path.join(root, fn), 'rb').read()).hexdigest() # get the hash value of the file
        filesInFolder_List =[[pathName,File_Path,File_Name,File_Extention,size,created_time,modified_time,userinfo,hash_md5]]
        filesInFolder_data1 = pd.DataFrame(filesInFolder_Listcolumns=file_Columns)
        filesInFolder_data = filesInFolder_data.append(filesInFolder_data1,ignore_index=True)

# Rename the file names of the files in folder if such file name matches with the files names in excel data
for i, f in enumerate(os.listdir(FolderPathToCheck)):
    src = os.path.join(FolderPathToCheck, f)
    for idx, row in excel_data.iterrows():
        if f == excel_data.loc[idx,'FileName']:
            dst = os.path.join(FolderPathToCheck, excel_data.loc[idx,'FileRenamed'])
            os.rename(src, dst)


# Split the "File_Path_Name" column into seperate columns as File Name, File Path and File extention        

for idx, row in filesInFolder_data.iterrows():
    filesInFolder_data.loc[idx,'File_Path'], filesInFolder_data.loc[idx,'File_Name'] = os.path.split(filesInFolder_data.loc[idx,'File_Path_Name'])
    filesInFolder_data.loc[idx,'File_Extention'] =  Path(filesInFolder_data.loc[idx,'File_Path_Name']).suffix 

# Write the details of files in the folder to a excel sheet.

writer = ExcelWriter(FolderPathToSaveResult+"/All_File_Header_Data.xlsx")
filesInFolder_data.to_excel(writer,'Sheet1',index=False)
writer.save()

# Write the details of files in the folder to a excel sheet.
writer = ExcelWriter(FolderPathToSaveResult+"/excel_data_new.xlsx")
excel_data.to_excel(writer,'Sheet1',index=False)
writer.save()
