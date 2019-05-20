# -*- coding: utf-8 -*-
"""
Created on Tue Mar 19 12:28:20 2019

@author: AYu5
"""

import sys
import pandas as pd
import numpy as np
import string
import re
from os import listdir
from os.path import isfile, join
from pandas import ExcelWriter



# Change following 3 variables into current positions and date

input_position = (r'C:\Users\ayu5\Desktop\left right program\0430LR\03data_test\Complete Parsed - To Be Reviewed')
output_position = (r'C:\Users\ayu5\Desktop\left right program\0430LR\output')
date =  '04182019'


input_position = input_position.replace('\\','/')
output_position = output_position.replace('\\','/')


function_list = ["Actor","Casting","Consultant","Director","Financier","GenericFunction",
                 "Producer","Researcher","Rights","TermDeal","Writer"]

files_list = [f for f in listdir(input_position) if isfile(join(input_position, f))]

for file in files_list:
    if "$" in file:
        files_list.remove(file)
        

#Except Rights & Termdeal
tab_list = ["Contingent_Compensation","Fixed_Compensation",#"Expenses","Location","Office_Secretary",
            #"On_Screen_Credit&Paid_Ad","Other_Terms","Transportation",
            "Payment"]

#Tabs for rights
rights_tab_list = ["Contingent_Comp_RightsIn","Contingent_Comp_RightsOut","Fixed_Comp_RightsIn&Out",#"Expenses",
                   #"Location","Office_Secretary","On_Screen_Credit&Paid_Ad","Other_Terms","Transportation",
                   "Payment"]

#Tabs for TermDeal
term_tab_list = ["Contingent_Compensation","Guarantee_Term_Deals","Overhead",#"Expenses","Location","Office_Secretary",
            #"On_Screen_Credit&Paid_Ad","Other_Terms","Transportation",
            "Payment"]

date_column_list = ["Compensation.Start_Date","Rights Start Date","Compensation Commitment Date",
                    "Start_Date","Expiry_Date","Purchase_Date","Reversion_Date",
                    "Date Paid","Commit_Date","Term_Start","Term_End","Deal_Date",
                    "PAYMENT_DATE"]

warning_list = []

#  Delete left columns, change right columns' names
def left_right(df_ori,tab,file_name):
    df = df_ori.copy()
    column_list = df.columns
    
    # For writers payment, all "_" are missing
    # if ("Writer" in file_name) & (tab == "Payment"):
    #    for column in column_list:
    #        df.rename(columns={column:column.replace(" ","_")}, inplace=True)
    #    column_list = df.columns
    
    for column in column_list:
        if "Right_" in column:
            original = column.replace("Right_","")
            
            
            # In some tabs, there is "Darts_Division" & "Right_DARTS_DIVISION"
            if (original == "DARTS_DIVISION") & (original not in column_list) & ("Darts_Division" in column_list):
                original = "Darts_Division"
                
            # In Term deal, there is "Deal_Date" & "Right_DEAL_DATE"
            if original == "DEAL_DATE":
                original = "Deal_Date"
                print(column, original)
                df.drop(labels=original, axis=1, inplace=True)
                #leave Deal_Date not upper
                df.rename(columns={column:original}, inplace=True)
                
                if original in date_column_list and not np.issubdtype(df[original], np.datetime64):
                    warning_list.append((file_name,tab,original))
                continue
            
            print(column, original)
            df.drop(labels=original, axis=1, inplace=True)
            df.rename(columns={column:original.upper()}, inplace=True)
            
            if original in date_column_list and not np.issubdtype(df[original], np.datetime64):
                warning_list.append((file_name,tab,original))
    return df


# Main function, input, process, output
def InChangeOut(excel,tab_list,file_name):
    df_dict = {}
    for tab in tab_list:
        df = pd.read_excel(excel,tab)
        
        # Changer Function
        df = left_right(df,tab,file_name)
        
        df_dict[tab] = df
        print(file_name, tab)
        
    function_name = ""
    for func in function_list:
        if func in file_name:
            function_name = func
    
    
    out = output_position + "/" + function_name +"_Function_FINAL.xlsx"
    print(out)
    with ExcelWriter(out) as writer:
        for tab in df_dict.keys():
            df_dict[tab].to_excel(writer,tab,index = False)
        writer.save()
        
        
        
# Main Iteration        
for file in files_list:
    print(file)
    
    if "Rights" in file: #For Rights
        excel = pd.ExcelFile(input_position + '/' + file)
        InChangeOut(excel,rights_tab_list,file)
        continue
        
    if "TermDeal" in file: #For TermDeal
        excel = pd.ExcelFile(input_position + '/' + file)
        InChangeOut(excel,term_tab_list,file)
        continue

    # For other functions
    excel = pd.ExcelFile(input_position + '/' + file)
    InChangeOut(excel,tab_list,file)


if not warning_list:
    print("No date type issues.")
else:
    print("Warning!!!")
    print("Warning!!!")
    print("Warning!!!")
    print("We get date type issues as " + str(warning_list))












