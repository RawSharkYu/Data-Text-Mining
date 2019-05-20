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

input_position = (r'C:\Users\ayu5\Desktop\left right program\0318Program\Complete Parsed - To Be Reviewed')
output_position = (r'C:\Users\ayu5\Desktop\left right program\0430LR\output')
date =  '04182019'



input_position = input_position.replace('\\','/')
output_position = output_position.replace('\\','/')

files_list = [f for f in listdir(input_position) if isfile(join(input_position, f))]

for file in files_list:
    if "$" in file:
        files_list.remove(file)
        

#Except Rights & Termdeal
tab_list = ["Contingent_Compensation","Fixed_Compensation"]
            #"Expenses","Location","Office_Secretary",
            #"On_Screen_Credit&Paid_Ad","Other_Terms","Transportation",
            #"Payment"]
           
    #Tabs for rights
rights_tab_list = ["Contingent_Comp_RightsIn","Contingent_Comp_RightsOut","Fixed_Comp_RightsIn&Out"]
                    #"Expenses",
                   #"Location","Office_Secretary","On_Screen_Credit&Paid_Ad","Other_Terms","Transportation",
                  # "Payment"]
                  

#Tabs for TermDeal
term_tab_list = ["Contingent_Compensation"]
            #"Guarantee_Term_Deals","Overhead","Expenses","Location","Office_Secretary",
            #"On_Screen_Credit&Paid_Ad","Other_Terms","Transportation",
            #"Payment"]
                


# column names that need to be datetime format
date_column_list = ["Compensation.Start_Date","Rights Start Date","Compensation Commitment Date",
                    "Start_Date","Expiry_Date","Purchase_Date","Reversion_Date",
                    "Date Paid","Commit_Date","Term_Start","Term_End","Deal_Date",
                    "PAYMENT_DATE"]

warning_list = []

# column names that need to be dropped and changed
con_col = ["Compensation.Bonus_Type","Compensation.Royalty_%","Compensation.PP_%","Compensation.PP_np/gp",
"Compensation.Prod_Bonus_Amount","Compensation.Deferment_Amount","Compensation.Oscar_Bonus_Amount",
"Compensation.Golden_Globe_Bonus_Amount","Compensation.On_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
"Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_ProRata",
"Compensation.Over_Budget_Penalty_%","Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
"Compensation.Box_Office_Relationship","Compensation.Box_Office_Type","Compensation.Parsed_sentence",
"Compensation.Bonus_Amount","Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
"Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index","Compensation.Bonus_Compensation_Type",
"TERM_DEAL_DURATION","#_Duration","Qualifier_Duration","Compensation.Sole_Shared","Compensation.Writing_Credit_np/gp",
"Compensation.Writing_Credit_%","Index","DARTS_DIVISION","COMPENSATION_ID","COMPENSATION_AMOUNT",
"COMPENSATION_DESC","COMPENSATION_TYPE","DEAL_ID","FUNCTION","PROJECT_ID"]

fix_col = ["Compensation.Applicable_ind","Compensation.Compensation_Type","Compensation.Service_Type",
"Compensation.Fee_Type","Compensation.Start_Date","Compensation.Duration_#","Compensation.Duration_Freq",
"Compensation.Rate","Compensation.Rate_Freq","Compensation.Payment_Type","Compensation.%_Total",
"Compensation.%_Amount","Compensation.Start_Qualifier","Compensation.Payment_Start_Condition",
"Compensation.Service_Start_Condition","Compensatipon.Total_Guaranteed_Commitment",
"Compensation.Writing_Step","Compensation.Writing_Step_Amount","Rights Start Condition",
"Rights Start Date","Compensation.Check_#","Compensation Commitment Date","# of Payments",
"Frequency","Start_Date","Expiry_Date","Purchase_Date","Reversion_Date","Date Paid",
"Index","DARTS_DIVISION","COMPENSATION_ID","COMPENSATION_AMOUNT","COMPENSATION_DESC",
"COMPENSATION_TYPE","DEAL_ID","FUNCTION","PROJECT_ID"]

column_names = {"Contingent_Compensation":con_col,"Fixed_Compensation":fix_col,
                "Fixed_Comp_RightsIn&Out":fix_col,"Contingent_Comp_RightsIn":con_col,
                "Contingent_Comp_RightsOut":con_col}


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

# Insert blank columns
def insert_blank(df_ori,tab,file_name):
    df = df_ori.copy()
    if tab not in column_names.keys():
        return df
    else:
        all_cols = column_names[tab]
        df_cols = df.columns
        for col in all_cols:
            if col not in df_cols:
                df.insert(loc = 0,column = col,value = np.NaN)
        df = df[all_cols]
    return df



def InChange(con_df,fix_df,input_position,tab_list,file_name):
    excel = pd.ExcelFile(input_position + '/' + file_name)
    df_dict = {}
    for tab in tab_list:
        df = pd.read_excel(excel,tab)
        
        # Changer Function
        df = left_right(df,tab,file_name)
        
        # Insert Blank Columns
        df = insert_blank(df,tab,file_name)
        
        df_dict[tab] = df
        print(file_name, tab)
    
    
    
    if "Rights" in file_name:
        df_dict["Contingent_Compensation"] = df_dict["Contingent_Comp_RightsIn"].append(df_dict["Contingent_Comp_RightsOut"])
        df_dict["Fixed_Compensation"] = df_dict["Fixed_Comp_RightsIn&Out"]
        df_dict.pop('Contingent_Comp_RightsIn', None)
        df_dict.pop('Contingent_Comp_RightsOut', None)
        df_dict.pop('Fixed_Comp_RightsIn&Out', None)
        
    
    if "Contingent_Compensation" in df_dict.keys():
        con_df = con_df.append(df_dict["Contingent_Compensation"])
        
    if "Fixed_Compensation" in df_dict.keys():
        fix_df = fix_df.append(df_dict["Fixed_Compensation"])        
         
    return con_df.reset_index(drop=True),fix_df.reset_index(drop=True)


# Create Empty DataFrame to store combined data
contin_df = pd.DataFrame(columns = con_col)
fixed_df = pd.DataFrame(columns = fix_col)



# Main iteration
for file in files_list:
    print(file)
    
    if "Rights" in file: #For Rights
        contin_df,fixed_df = InChange(contin_df,fixed_df,input_position,rights_tab_list,file)
        continue
        
    if "TermDeal" in file: #For TermDeal
        contin_df,fixed_df = InChange(contin_df,fixed_df,input_position,term_tab_list,file)
        continue

    # For other functions
    contin_df,fixed_df = InChange(contin_df,fixed_df,input_position,tab_list,file)


#Output
contin_df.at[contin_df[['DARTS_DIVISION','COMPENSATION_DESC','COMPENSATION_ID']].duplicated(),'COMPENSATION_DESC'] = np.NaN
fixed_df.at[fixed_df[['DARTS_DIVISION','COMPENSATION_DESC','COMPENSATION_ID']].duplicated(),'COMPENSATION_DESC'] = np.NaN

with ExcelWriter(output_position + "/" + "Contingent_Appended.xlsx") as writer:
    contin_df.to_excel(writer,"Contingent_Compensation",index = False)
    writer.save()
    
with ExcelWriter(output_position + "/" + "Fixed_Appended.xlsx") as writer:
    fixed_df.to_excel(writer,"Fixed_Compensation",index = False)
    writer.save()
    

if not warning_list:
    print("No date type issues.")
else:
    print("Warning!!!")
    print("Warning!!!")
    print("Warning!!!")
    print("We get date type issues as " + str(warning_list))
    
    