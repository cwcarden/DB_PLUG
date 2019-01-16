import pyodbc
import csv
import os
import pandas as pd
import time

'''
Dependencies are as follows: Use Pip install
- pyodbc
- csv (built in to python)
- pandas
- xlrd (pluggin for xlsx files)
- XlsxWriter
'''
#Prepare import data
#1. rename import xlsx sheet tab to "Sheet1"
#2. rename xlsx grow plan to "plan.xlsx"

#Step 1 build a DataFrame from xlsx file and write to a csv
def panda():
    #Assign spreadsheet this name
    file = 'plan.xlsx'
    #Load Spreadhseet
    xl = pd.ExcelFile(file)
    #print the sheet names
    print(xl.sheet_names)
    #Load a sheet into a DataFrame by name: df1
    df1 = xl.parse('Sheet1')
    writer = pd.ExcelWriter('plan.xlsx', engine='xlsxwriter')
    df1.drop(df1.index[:4], inplace=True)
    df1.rename(
        columns={
            "Unnamed: 1":"Area",
            "Unnamed: 4":"Hybrid",
            "Unnamed: 8":"Certified",
            "Unnamed: 13":"Total 50LB units",
            "Unnamed: 17":"Female Acres",
            "Unnamed: 18":"Units/Female Acre 50lb",
            "Unnamed: 19":"Units/GA",
            "Unnamed: 21":"%F",
            "Unnamed: 22":"Gross Acres",
            "Unnamed: 28":"Female Acres",
            "Unnamed: 30":"Female Parent",
            "Unnamed: 37":"Male Inbred",
            }, inplace=True)
    df1.to_csv('new.csv')

def clean_up():
    file = 'new.csv'
    df2 = pd.read_csv(file)
    df2.drop(['Tentative - CY19 Sorghum Acreage Plan', 'Unnamed: 2', 
    'Unnamed: 3', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 9', 
    'Unnamed: 10', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 14', 'Unnamed: 15',
    'Unnamed: 16', 'Unnamed: 20', 'Unnamed: 23', 'Unnamed: 24', 'Unnamed: 25', 
    'Unnamed: 26', 'Unnamed: 27', 'Unnamed: 29', 'Unnamed: 31', 'Unnamed: 32', 
    'Unnamed: 33', 'Unnamed: 34', 'Unnamed: 35', 'Unnamed: 36', 'Unnamed: 38', 
    'Unnamed: 39', 'Unnamed: 40', 'Unnamed: 41', 'Unnamed: 42', 'Unnamed: 43', 
    'Unnamed: 44', 'Unnamed: 45' ], axis=1).to_csv('new2.csv')


#Connection to microsoft access database
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\cardench\Desktop\db_plug\sorghum2018.accdb;'
)

#View to display all table data
def all_data():
    cnxn = pyodbc.connect(conn_str)
    crsr = cnxn.cursor()
    crsr.execute('SELECT Hybrid, Area, [Units/GA], [Female Acres], [50 LB Units], [Female Parent], [Male Parent], [Gross Acres], [Units/Female Acre 50lb], Certified, [%F] FROM [Budget Acreage Plan]')
    data = crsr.fetchall()
    print(data)

#First delete all rows in table Budget Acreage Plan
def delete_rows():
    cnxn = pyodbc.connect(conn_str)
    crsr = cnxn.cursor()
    crsr.execute('DELETE FROM [Budget Acreage Plan]')
    crsr.commit()

#Second read new budget plan from csv and clean up for importing 
