# -*- coding: utf-8 -*-
"""
Created on Mon May  6 15:46:02 2019

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



date =  '0506'


dict_combined = {}

function_list = ["Actor","Casting","Consultant","Director","Financier","Generic_Functions",
                 "Producer","Researcher","Rights","Termdeal","Writer"]

files_list = [f for f in listdir(input_position) if isfile(join(input_position, f)) and "Screen" in f]

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
                        'FUNCTION':'Right_FUNCTION','ON_SCREEN_CREDIT':'Right_ON_SCREEN_CREDIT',
                        'PAID_AD':'Right_PAID_AD','PROJECT_ID':'Right_PROJECT_ID'})
    

    combined = pd.concat([existed_df,new_df],sort=False,ignore_index=True)
    
    dict_combined[func] = combined
    
    
    
# 1.line spliter
def line_spliter(orig_df,index,num):
    df = orig_df.copy()
    row = df.loc[index]
    print(num - 1)
    for j in range(num - 1):
        df = df.append(row)
    return df



# shared or sep card/seperated/single
def card_paster(df, index):
    
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    
    screen = row["Right_ON_SCREEN_CREDIT"]
    
    if not pd.isnull(screen) and re.findall(r"(?i)shared", screen):
        df.at[index,"Credit.Card"] = "Shared"
    if not pd.isnull(screen) and re.findall(r"(?i)(sep card|separate card|single)", screen):
        df.at[index,"Credit.Card"] = "Single"
    
    return df  


def position_regex(test_str):
    if pd.isnull(test_str):
        return np.nan
    
    if re.findall(r"(?i)(1st|first) posi?t?i?o?n?", test_str):
        return "1st Position"
    if re.findall(r"(?i)(2nd|second) posi?t?i?o?n?", test_str):
        return "2nd Position"
    if re.findall(r"(?i)(3rd|third) posi?t?i?o?n?", test_str):
        return "3rd Position"
    if re.findall(r"(?i)(4th|fourth) posi?t?i?o?n?", test_str):
        return "4th Position"
    if re.findall(r"(?i)(5th|fifth) posi?t?i?o?n?", test_str):
        return "5th Position"
    if re.findall(r"(?i)(6th|sixth) posi?t?i?o?n?", test_str):
        return "6th Position"
    if re.findall(r"(?i)(7th|seventh) posi?t?i?o?n?", test_str):
        return "7th Position"
    if re.findall(r"(?i)(8th|eighth) posi?t?i?o?n?", test_str):
        return "8th Position"
    if re.findall(r"(?i)last posi?t?i?o?n?", test_str):
        return "Last Position"

    return np.nan


# For Credit.Position
def position_paster(df, index):
    
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    
    screen = row["Right_ON_SCREEN_CREDIT"]
    paid = row["Right_PAID_AD"]
    
    reg_1 = position_regex(paid)
    reg_2 = position_regex(screen)
    
    if not pd.isnull(reg_1):
        df.at[index, "Credit.Position"] = reg_1
    
    if not pd.isnull(reg_2):
        df.at[index, "Credit.Position"] = reg_2
    
    return df


# For Credit.Paid_Ad
def paid_ad_paster(df, index):
    
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    
    paid = row["Right_PAID_AD"]
    
    if not pd.isnull(paid):
        df.at[index, "Credit.Paid_Ad"] = "TRUE"
    else:
        df.at[index, "Credit.Paid_Ad"] = "FALSE"
    
    return df

# For Credit.Billing_Block
def billing_block_paster(df, index):
    
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    
    paid = row["Right_PAID_AD"]
    
    if not pd.isnull(paid) and re.findall(r"(?i)billing block", paid):
        df.at[index,"Credit.Billing_Block"] = "Billing Block"
    
    return df


# if in Right_ON_SCREEN_CREDIT or Right_PAID_AD there is "bug" shown
def bug_logo_paster(df, index):
    
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    
    screen = row["Right_ON_SCREEN_CREDIT"]
    paid = row["Right_PAID_AD"]
    
    reg_1 = True if not pd.isnull(paid) and re.findall(r"(?i)\Wbug\W", paid) else False
    reg_2 = True if not pd.isnull(screen) and re.findall(r"(?i)\Wbug\W", screen) else False
    
    if reg_1 or reg_2:
        df.at[index, "Credit.Bug_Logo"] = "Bug Logo"

    return df

# For Credit.Animated_Logo, not spliting
def animated_logo_paster(df, index):
    
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    
    screen = row["Right_ON_SCREEN_CREDIT"]
    paid = row["Right_PAID_AD"]
    
    reg_1 = True if not pd.isnull(paid) and re.findall(r"(?i)\Wanimated\W", paid) else False
    reg_2 = True if not pd.isnull(screen) and re.findall(r"(?i)\Wanimated\W", screen) else False
    
    if reg_1 or reg_2:
        df.at[index, "Credit.Animated_Logo"] = "Animated Logo"
    
    return df     



# Main Title, M/T
# End Title, E/T
# above/before title
# after title

def main_end_title(df, index):
    
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    
    screen = row["Right_ON_SCREEN_CREDIT"]
    
    if not pd.isnull(screen) and re.findall(r"(?i)(main title|M\/T)", screen):
        df.at[index,"Credit.Main_Title_End_Title"] = "Main Title"
        
    if not pd.isnull(screen) and re.findall(r"(?i)(end title|E\/T)", screen):
        df.at[index,"Credit.Main_Title_End_Title"] = "End Title"
        
    if not pd.isnull(screen) and re.findall(r"(?i)(above title|before title|above\/before title)", screen):
        df.at[index,"Credit.Main_Title_End_Title"] = "Above/Before title"

    if not pd.isnull(screen) and re.findall(r"(?i)(after title)", screen):
        df.at[index,"Credit.Main_Title_End_Title"] = "After Title"
        
    return df



def type_credit_regex(test_str,function):
    
    type_list = []
    
    # "Actor"
    # see function if talent
    if function == "Talent":
        type_list.append("Actor")
        
    if pd.isnull(test_str):
        return type_list
    
    # "Associate Producer Credit"
    if re.findall(r"(?i)(Associ?a?t?e? Produ?c?e?r?|\WAP\W)", test_str):
        type_list.append("Associate Producer Credit")
    
    # "Based On"
    if re.findall(r"(?i)Based (on|upon)", test_str):
        type_list.append("Based On")
    
    # "Casting By"
    if re.findall(r"(?i)Casting by", test_str):
        type_list.append("Casting By")
    
    # "Consultant"
    if re.findall(r"(?i)consult", test_str):
        type_list.append("Consultant")
    
    
    # "Co-Producer Credit" 
    if re.findall(r"(?i)co-produ?c?e?r?", test_str):
        type_list.append("Co-Producer Credit")
    
    # "Directed by"
    if re.findall(r"(?i)directed by", test_str):
        type_list.append("Directed by")
    
    # "Exectuive Producer"
    if re.findall(r"(?i)(\WEP\W|Execu?t?i?v?e? Produ?c?e?r?)", test_str):
        type_list.append("Executive Producer")
    
    # "Film By"
    if re.findall(r"(?i)Film By", test_str):
        type_list.append("Film By")
    
    # "Per DGA"
    if re.findall(r"(?i)Per DGA", test_str):
        type_list.append("Per DGA")
    
    # "Per WGA"
    if re.findall(r"(?i)Per WGA", test_str):
        type_list.append("Per WGA")
    
    # "Produced By"
    if re.findall(r"(?i)Produced by", test_str):
        type_list.append("Produced By")

    # "Production Company Credit"
    if re.findall(r"(?i)Produ?c?t?i?o?n? com?p?a?n?y?", test_str):
        type_list.append("Production Company Credit")

    # "Production"
    if re.findall(r"(?i)Produ?c?t?i?o?n Credit", test_str):
        type_list.append("Production")
    
    return type_list


# "Credit.Type_of_Credit"
def type_credit_split(df,index):
    row = df[df["fake_id"] == index]
    ind = row.index[0]
    row = df.loc[ind]
    
    screen = row["Right_ON_SCREEN_CREDIT"]
    function = row["Right_FUNCTION"]
    
    type_list = type_credit_regex(screen,function)
    
    num = len(type_list)
    
    if num > 0:
        if num == 1:
            df.at[ind,"Credit.Type_of_Credit"] = type_list[0]
        else:
            # Split if there is more than 1 type.
            df = line_spliter(df, ind, num)
            df.at[ind,"Credit.Type_of_Credit"] = type_list
    
    return df.reset_index(drop=True) 

# All has same order of columns
column_order = ["Darts_Division","DEAL_ID","FUNCTION","PROJECT_ID","ON_SCREEN_CREDIT","PAID_AD","Credit.Type_of_Credit",
                "Credit.Main_Title_End_Title","Credit.Card","Credit.Position","Credit.Paid_Ad","Credit.Bug_Logo",
                "Credit.Animated_Logo","Credit.Billing_Block","Index","Right_DARTS_DIVISION","Right_DEAL_ID",
                "Right_FUNCTION","Right_ON_SCREEN_CREDIT","Right_PAID_AD","Right_PROJECT_ID"]


for func in dict_combined.keys():
    
    df = dict_combined[func]
    # To iteration, must using an one-to-one unchangable id column
    df["fake_id"] = df.index
    id_list = df.loc[df["Index"] == "New"]["fake_id"].unique()
    
    for ind in id_list:
        print(func)
        print(ind)
        df = card_paster(df, ind)
        df = position_paster(df, ind)
        df = paid_ad_paster(df, ind)
        df = billing_block_paster(df, ind)
        df = bug_logo_paster(df, ind)
        df = animated_logo_paster(df, ind)
        df = main_end_title(df, ind)
        df = type_credit_split(df, ind)
    
    #One more step, choose excat columns with exact order.
    df = df[column_order]
    
    with pd.ExcelWriter(output_position + '/' + date + "_OnScreenCredit_" + func + "_" +
                              "Parsed.xlsx", engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='OnScreenCredit',index = False)




