# -*- coding: utf-8 -*-
"""
Created on Thu May  2 10:33:37 2019

@author: AYu5
"""

import pandas as pd
import re
import numpy as np
from os import listdir
from os.path import isfile, join
import sys
from decimal import Decimal
import datetime as dt


input_position = (r'C:\Users\ayu5\Desktop\Old_Data_Parsing\data_parsing_0128\data_parsing_raw\Use Every Time_Output Data - Copy')
output_position = (r'C:\Users\ayu5\Desktop\PythonForDP\output')
input_position = input_position.replace('\\','/')
output_position = output_position.replace('\\','/')

date =  '0502'



# Don't change any lines below without understanding.

dict_combined = {}

function_list = ["Actor","Casting","Consultant","Director","Financier","Generic_Functions",
                 "Producer","Researcher","Rights","Termdeal","Writer"]


# Read file names

files_list = [f for f in listdir(input_position) if isfile(join(input_position, f)) and "Other" in f]

for file in files_list:
    if "$" in file:
        files_list.remove(file)
        
for func in function_list:
    
    print(func)
    
    existed_file = [f for f in files_list if func in f and "Existed" in f][0]
    new_file = [f for f in files_list if func in f and "New" in f][0]
    existed_df = pd.read_csv(input_position + '/' + existed_file)
    new_df = pd.read_csv(input_position + '/' + new_file)

    new_df = new_df.rename(index=str, 
               columns={'DARTS_DIVISION':'Right_DARTS_DIVISION','DEAL_ID':'Right_DEAL_ID',
                        'FUNCTION':'Right_FUNCTION','OTHER_TERMS':'Right_OTHER_TERMS',
                        'PAY_OR_PLAY_DESC':'Right_PAY_OR_PLAY_DESC','PROGRESS_TO_PROD':'Right_PROGRESS_TO_PROD',
                        'PROJECT_ID':'Right_PROJECT_ID'})
    

    combined = pd.concat([existed_df,new_df],sort=False,ignore_index=True)
    
    dict_combined[func] = combined
    
# General function: line spliter
def line_spliter(orig_df,index,num):
    df = orig_df.copy()
    row = df.loc[index]
    print(num - 1)
    for j in range(num - 1):
        df = df.append(row)
    return df


# Only change 1 column: Other_Terms.Other_Terms_Type
# Look into 3 columns Right_OTHER_TERMS,Right_PAY_OR_PLAY_DESC,Right_PROGRESS_TO_PROD
# Cuts & Previews, Designations, First Negotiation, Pay or Play, Progress to Production, Sequels/Remakes
def type_spliter(df, index):
    # For now, only 1 row
    
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    print(ind)
    
    other_terms = row['Right_OTHER_TERMS']
    pay_or_play = row['Right_PAY_OR_PLAY_DESC']
    prog_to_prod = row['Right_PROGRESS_TO_PROD']
    
    print(other_terms)
    if pd.isnull(other_terms):
        cut_and_pre, first_nego, design, sequ_rema = False, False, False, False
    else:
        
        cut_and_pre = True if len(re.findall(r"(?i)cut.*preview|preview.*cut", other_terms)) > 0 else False    
        first_nego = True if len(re.findall(r"(?i)(first|1st) negotia", other_terms)) > 0 else False
        
        design = True if len(re.findall(r"(?i)notices\:|pa?yme?nts to\:|checks go to\:|contact info\:", other_terms))\
        > 0 else False
        
        sequ_rema = True if len(re.findall(r"(?i)seq.*remake", other_terms)) > 0 else False
    pay_or_play = (not pd.isnull(pay_or_play))
    prog_to_pord = (not pd.isnull(prog_to_prod))
    
    
    other_terms_type = []
    
    if cut_and_pre:
        other_terms_type.append("Cuts & Previews")
    if first_nego:
        other_terms_type.append("First Negotiation")
    if design:
        other_terms_type.append("Designations")
    if sequ_rema:
        other_terms_type.append("Sequels/Remakes")
    if pay_or_play:
        other_terms_type.append("Pay or Play")
    if prog_to_pord:
        other_terms_type.append("Progress to Production")
    
    num = len(other_terms_type)
    
    if num > 1:
        df = line_spliter(df, ind, num)
        df.at[ind,"Other_Terms.Other_Terms_Type"] = other_terms_type
    elif num == 1:
        df.at[ind,"Other_Terms.Other_Terms_Type"] = other_terms_type[0]
    else:
        return df
    
    
    return df.reset_index(drop=True)

# All has same order of columns
column_order = ["Darts_Division","DEAL_ID","FUNCTION","OTHER_TERMS","PROJECT_ID","original_OTHER_TERMS",
                              "PAY_OR_PLAY_DESC","PROGRESS_TO_PROD","Other_Terms.Other_Terms_Type","Index",
                              "Right_DARTS_DIVISION","Right_DEAL_ID","Right_FUNCTION","Right_OTHER_TERMS",
                              "Right_PAY_OR_PLAY_DESC","Right_PROGRESS_TO_PROD","Right_PROJECT_ID"]



# Iteration and output
for func in dict_combined.keys():
    
    df = dict_combined[func]
    # To iteration, must using an one-to-one unchangable id column
    df["fake_id"] = df.index
    id_list = df.loc[df["Index"] == "New"]["fake_id"].unique()
    
    for ind in id_list:
        print(func)
        print(ind)
        df = type_spliter(df, ind)
    
    #One more step, choose excat columns with exact order.
    df = df[column_order]
    
    with pd.ExcelWriter(output_position + '/' + date + "_Other_Terms_" + func + "_" +
                              "Parsed.xlsx", engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Other_Terms',index = False)