# -*- coding: utf-8 -*-
"""
Created on Wed May 15 09:50:10 2019

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
import locale



input_position = (r'C:\Users\ayu5\Desktop\PythonForDP\04ConData\Contingent')
output_position = (r'C:\Users\ayu5\Desktop\PythonForDP\output')
input_position = input_position.replace('\\','/')
output_position = output_position.replace('\\','/')

date =  '0515'


# Don't change any line below if not certain. 



# Inputing files

function_name_list = ['Actor','Casting','Consultant','Director','Financier','Termdeal',
                 'RightsIn','RightsOut','Generic_Functions','Producer','Writer']

def existed_change(existed_fix):
    df = existed_fix.copy()
    if 'Index1' in df.columns:
        df = df.drop(columns=['Index1'])
    return df

def new_change(new_fix):
    df = new_fix.copy()
    df = df.rename(index=str, 
                   columns={'Darts_Division':'Right_DARTS_DIVISION','COMPENSATION_ID':'Right_COMPENSATION_ID',
                           'COMPENSATION_AMOUNT':'Right_COMPENSATION_AMOUNT','COMPENSATION_DESC':'Right_COMPENSATION_DESC',
                           'COMPENSATION_TYPE':'Right_COMPENSATION_TYPE','DEAL_ID':'Right_DEAL_ID',
                           'FUNCTION':'Right_FUNCTION','PROJECT_ID':'Right_PROJECT_ID'})
    return df



dict_combined = {}

for func in function_name_list:
    
    if func == 'RightsIn':
        existed_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_Rights_Cont_RightsIn_Existed.xlsx")
        new_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_Rights_Cont_RightsIn_New.xlsx")
        existed_df = pd.read_excel(existed_file,"Sheet1")
        new_df = pd.read_excel(new_file,"Sheet1")
        
    elif func == 'RightsOut':
        existed_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_Rights_Cont_RightsOut_Existed.xlsx")
        new_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_Rights_Cont_RightsOut_New.xlsx")
        existed_df = pd.read_excel(existed_file,"Sheet1")
        new_df = pd.read_excel(new_file,"Sheet1")
    
    else:
        existed_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_" + func + "_Cont_Existed.xlsx")
        new_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_" + func + "_Cont_New.xlsx")
        existed_df = pd.read_excel(existed_file,"Sheet1")
        new_df = pd.read_excel(new_file,"Sheet1")
    
    combined = pd.concat([existed_change(existed_df),new_change(new_df)],sort=False,ignore_index=True)
    
    dict_combined[func] = combined
    
#General Function Preparation
# Split one line (specified by index) into multiple lines.
# If there are more than 1 line with same index, this function 
# will split all the lines.
def line_spliter(orig_df,index,num):
    df = orig_df.copy()
    row = df.loc[index]
    print(num - 1)
    for j in range(num - 1):
        df = df.append(row)
    return df


# Transfer money string format from "$1.2K" into "1200"
def money_transfer(money_str):
    num = Decimal(re.sub(r'[^\d.]', '', money_str))
    if 'k' in money_str or 'K' in money_str:
        num = num * 1000
    if 'm' in money_str or 'M' in money_str:
        num = num * 1000000
    if 'b' in money_str or 'B' in money_str:
        num = num * 1000000000    
    return num

# Transfer box office money string format from "325000000" into "325M" 
# num should be a decimal 
def bo_money_back_transfer(num):
    integer = num.to_integral_value()
    length = len(str(integer))
    if length >= 10:
        return str(integer/1000000000) + 'B'
    if length >= 7:
        return str(integer/1000000) + 'M'
    if length >= 4:
        return str(integer/1000) + 'K'    
    
    return str(integer)

# Transfer bonus money string format from "1200" into "$1.2K" 
# num should be a decimal 
def bonus_money_back_transfer(num):
    locale.setlocale(locale.LC_ALL, '' )
    return locale.currency(num, grouping=True )


#General Class Preparation
class indicator:
    def __init__(self, indi, indi_type, position_start, position_end, desc):
        self.indi = indi
        self.indi_type = indi_type
        self.position_start = position_start
        self.position_end = position_end
        self.desc = desc
        
    def left(self, another):
        if self.position_end < another.position_start:
            return True
        else:
            return False
        
    def right(self, another):
        if self.position_start > another.position_end:
            return True
        else:
            return False        

    def between_indi(self, other_1, other_2):
        if self.right(self,other_1) and self.left(self,other_2):
            return True
        else:
            return False  
    
    def inside_block(self, block):
        if block.position_start <= self.position_start and self.position_end <= block.position_end:
            return True
        else:
            return False
        
class indi_list:
    def __init__(self, list_of_indi, indi_type, desc):
        self.list_of_indi = list_of_indi
        self.indi_type = indi_type
        self.length = len(list_of_indi)
        self.list = [indicator.indi for indicator in list_of_indi]
        self.set = set(self.list)
        self.desc = desc
        
    def add_at_left(self, indicator):
        self.list_of_indi = [indicator] + self.list_of_indi
        self.length = self.length + 1
        self.list = [indicator.indi] + self.list
        self.set = set(self.list)
        

# block means the string between 2 indicators. 
class block:
    def __init__(self, left_indi, right_indi, desc):
        self.left_indi = left_indi
        self.right_indi = right_indi
        self.position_start = left_indi.position_end
        self.position_end = right_indi.position_start
        self.desc = desc
        
    def left(self, another):
        if self.position_end < another.position_start:
            return True
        else:
            return False
        
    def right(self, another):
        if self.position_start > another.position_end:
            return True
        else:
            return False        

    def between_indi(self, other_1, other_2):
        if self.right(self,other_1) and self.left(self,other_2):
            return True
        else:
            return False  
    
    def inside_block(self, indi):
        if self.position_start <= indi.position_start and indi.position_end <= self.position_end:
            return True
        else:
            return False        
        
        
# Data Parsing starts from here
            
# First part is no need to split line
def compensation_type_paster(df, comp_id):
    row = df[df["Right_COMPENSATION_ID"] == comp_id]
    ind = row.index[0]
    df.at[ind, "Compensation.Bonus_Compensation_Type"] = df.loc[ind, 'Right_COMPENSATION_TYPE']
    return df

def writer_bonus_type_paster(df, comp_id):
    row = df[df["Right_COMPENSATION_ID"] == comp_id]
    ind = row.index[0]
    df.at[ind, "Compensation.Bonus_Type"] = "Writing Credit Bonus"
    return df

# No need to split line
# For Compensation.PP_np/gp and Compensation.PP_%
# Except for Writer and Director, which two have "Shared Credit Bonus" and "Sole Credit Bonus"
def pp_np_num_paster(df, comp_id):
    row = df[df["Right_COMPENSATION_ID"] == comp_id]
    ind = row.index[0]
    row = df.loc[ind]
    desc = row["Right_COMPENSATION_DESC"]
    
    
    result = re.findall(r"(?i)(\d\.?\d?)\%\s(GP|NP)",desc)
    
    #5% GP
    if len(result) == 1:
        df.at[ind, "Compensation.PP_np/gp"] = result[0][1].replace("NP","np").replace("GP",'gp')
        df.at[ind, "Compensation.PP_%"] = result[0][0]
        df.at[ind, "Compensation.Bonus_Type"] = "Percentage Participation"
        return df,True
        
    #5% GP, 3% GP -------> only GP
    elif len(result) > 1 & len(set([x[1] for x in result])) == 1:
        df.at[ind, "Compensation.PP_np/gp"] = result[0][1].replace("NP","np").replace("GP",'gp')
        df.at[ind, "Compensation.Bonus_Type"] = "Percentage Participation"
        return df,True

    return df,False

# For writer only
# Maybe need to split line
# For Compensation.PP_np/gp and Compensation.PP_%
def writer_pp_np_num_spliter(df, comp_id):
    #at now, only 1 row
    row = df[df["Right_COMPENSATION_ID"] == comp_id]
    ind = row.index[0]
    row = df.loc[ind]
    desc = row["Right_COMPENSATION_DESC"]
    
    np_gp = re.findall(r"(?i)(GP|NP)",desc)
    
    percentage = re.findall(r"(?i)(\d\.?\d?)\%",desc)
    
    sole_shared = re.findall(r"(\sshared\s.*\ssole\s|\ssole\s.*\sshared\s)",desc)   
    
    if len(set([i.lower() for i  in np_gp])) == 1:
        if len(sole_shared) == 1 and len(percentage) == 2 and  Decimal(percentage[0])/Decimal(percentage[1]) == 2:
         
            df = line_spliter(df, ind, 2)


            df.at[ind, "Compensation.Writing_Credit_np/gp"] = np.repeat(np_gp[0].lower(),2)

            df.at[ind, "Compensation.Writing_Credit_%"] = percentage
            df.at[ind, "Compensation.Sole_Shared"] = ["Sole Directing Credit", "Shared Directing Credit"]
            return df.reset_index(drop=True),True
        elif len(sole_shared) == 0:
            if len(percentage) == 1:
                df.at[ind, "Compensation.Writing_Credit_np/gp"] = np_gp[0].lower()
                df.at[ind, "Compensation.Writing_Credit_%"] = percentage[0]
                return df,True
            elif len(percentage) > 1:
                df.at[ind, "Compensation.Writing_Credit_np/gp"] = np_gp[0].lower()
                return df,True
            
    return df,False



# Regex to catch indicators
def box_office_regex(str_test):

    catching_list = []
    
    p = re.compile(r"(?i)(\d{1,3}\.?\d*M|\d{1}\.?\d*B)")
    
    for m in p.finditer(str_test):
        catching_list.append(indicator(money_transfer(m.group()),"box_office",m.start(),m.end(),str_test))
        
    return indi_list(catching_list,"box_office",str_test)

def bonus_regex(str_test):
    
    catching_list = []
    
    p = re.compile(r"(?i)(\d{1,3}\,\d{3}|\d{1,3}k)")
    
    for m in p.finditer(str_test):
        catching_list.append(indicator(money_transfer(m.group()),"bonus",m.start(),m.end(),str_test))
        
    return indi_list(catching_list,"bonus",str_test)

def bo_type_regex(str_test):
    
    catching_list = []
    
    p = re.compile(r"(?i)(DBO|WWBO)")
    
    for m in p.finditer(str_test):
        catching_list.append(indicator(m.group().upper(),"box_office_type",m.start(),m.end(),str_test))
   
    return indi_list(catching_list,"box_office_type",str_test)

def or_regex(str_test):   
    
    catching_list = []
    
    p = re.compile(r"\sor\s")
    
    for m in p.finditer(str_test):
        catching_list.append(indicator(m.group().lower(),"or",m.start(),m.end(),str_test))
   
    return indi_list(catching_list,"or",str_test)

#For the show of "@ ea of" or "@ earlier of".
def at_ea_of_regex(str_test):

    catching_list = []
    
    p = re.compile(r"(?i)(\@ ea of|\@ earlier of)")
    
    for m in p.finditer(str_test):
        catching_list.append(indicator(m.group().lower(),"at each of",m.start(),m.end(),str_test))
   
    return indi_list(catching_list,"at each of",str_test)

#For only Writer and Director.
    
# 1/2 the amounts.
# 1/2 for shared credit
    
def share_sole_credit_regex(str_test):
    if re.findall(r"(?i)(1\/2 the amount|1\/2 for shared)", str_test):
        return True 
    else:
        return False
    
def share_credit_regex(str_test):
    catching_list = []
    
    p = re.compile(r"\sshared\s")
    
    for m in p.finditer(str_test):
        catching_list.append(indicator(m.group().lower(),"shared",m.start(),m.end(),str_test))
   
    return indi_list(catching_list,"shared",str_test)    

def sole_credit_regex(str_test):
    catching_list = []
    
    p = re.compile(r"\ssole\s")
    
    for m in p.finditer(str_test):
        catching_list.append(indicator(m.group().lower(),"sole",m.start(),m.end(),str_test))
   
    return indi_list(catching_list,"sole",str_test)      


# There are several patterns to split
# First see if the data is Director, Writer or not. ( if Director or Writer, we need to see if "shared credit")
    
def type_spliter_1(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list):
    print("This is type 1!")
    num = box_office_list.length
    
    df = line_spliter(df, ind, num)
    
    df.at[ind,"Compensation.Box_Office_Type"] = np.repeat(bo_type_list.list[0], num)
    df.at[ind,"Compensation.Box_Office_Qualifier2"] = np.repeat('or', num)
    df.at[ind,"Compensation.Bonus_Amount"] = bonus_list.list
    df.at[ind,"Compensation.Box_Office_Amount"] = box_office_list.list
    df.at[ind,"Compensation.Box_Office_Qualifier"] = np.repeat('Equal or Greater than', num)
    df.at[ind,"Compensation.Box_Office_Index"] = [i for i in range(1,num+1)]
    df.at[ind,"Compensation.Box_Office_Relationship"] = "("+") AND (".join([str(i) for i in range(1,num+1)]) + ")"
    df.at[ind,"Compensation.Parsed_sentence"] =  [bo_money_back_transfer(m) + '  ' + bonus_money_back_transfer(n) \
                                              for m,n in zip(box_office_list.list,bonus_list.list)]    
    
    return df
    
def type_spliter_2(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list):
    print("This is type 2!")
    num = box_office_list.length
    df = line_spliter(df, ind, num)

    
    # This may be able to modified
    type_list = []
    for i in range(bonus_list.length):
        type_list.append("DBO")
        if i < or_list.length:
            type_list.append("WWBO")
            
    df.at[ind,"Compensation.Box_Office_Type"] = type_list

    df.at[ind,"Compensation.Box_Office_Qualifier2"] = np.repeat('or', num)


    # For "Compensation.Bonus_Amount"
    bonus_list_all = []
    for i in range(bonus_list.length):
        bonus_list_all.append(bonus_list.list[i])
        if i < or_list.length:
            bonus_list_all.append(bonus_list.list[i])
    df.at[ind,"Compensation.Bonus_Amount"] = bonus_list_all            

    df.at[ind,"Compensation.Box_Office_Amount"] = box_office_list.list

    df.at[ind,"Compensation.Box_Office_Qualifier"] = np.repeat('Equal or Greater than', num)

    df.at[ind,"Compensation.Box_Office_Index"] = [i for i in range(1,num+1)]


    # For "Compensation.Box_Office_Relationship"
    relationship = "("
    i = 1
    while i < num+1:
        relationship = relationship + str(i)
        if i <= 2*or_list.length:
            i = i + 1
            relationship = relationship + " OR " + str(i)
        if i == num:
            break
        relationship = relationship + ") AND ("
        i = i + 1
    relationship = relationship + ")"

    df.at[ind,"Compensation.Box_Office_Relationship"] = relationship

    df.at[ind,"Compensation.Parsed_sentence"] =  [bo_money_back_transfer(m) + '  ' + bonus_money_back_transfer(n) \
                                                  for m,n in zip(box_office_list.list,bonus_list_all)]  
    
    
    
    return df    

def type_spliter_3(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list):
    print("This is type 3!")
    num = box_office_list.length
    desc = df.loc[ind,"Right_COMPENSATION_DESC"]
    print("num is " + str(num))
    right_comp = df.loc[ind,"Right_COMPENSATION_AMOUNT"]
    df = line_spliter(df, ind, num)
    df.at[ind,"Compensation.Box_Office_Qualifier2"] = np.repeat('or', num)
    df.at[ind,"Compensation.Box_Office_Qualifier"] = np.repeat('Equal or Greater than', num)
    df.at[ind,"Compensation.Box_Office_Index"] = [i for i in range(1,num+1)]
    df.at[ind,"Compensation.Box_Office_Amount"] = box_office_list.list

    
    # indicator(self, indi, indi_type, position_start, position_end, desc):
    if at_ea_list.list_of_indi[0].position_start == 0:
        # for those start with "@ ea of"
        bonus_list = bonus_list.add_at_left(indicator(str(right_comp),'bonus',-2,-1,desc))

    # CUT by bonus amount 
    bonus_list_all = []
    bo_type_list_all = []
    relation_string = ""
    
    block_list = []
    
    
    
    # block(self, left_indi, right_indi, desc):
    for i in range(bonus_list.length - 1):
        block_list.append(block(bonus_list.list_of_indi[i], bonus_list.list_of_indi[i + 1], desc))
        
    # Add the tail, not including head
    block_list.append(block(bonus_list.list_of_indi[-1],
                            indicator("End", "End", len(desc), len(desc), desc), desc))
            

    # number of blocks equals number of bonus indicators


    # check if there is any "@ ea of" inside each block
    for i in range(len(block_list)):
        
        # start_num is used for relation_string
        start_num = len(bonus_list_all)

        bonus = bonus_list.list_of_indi[i]
        bouns_position = (bonus.position_start, bonus.position_end)
        block_current = block_list[i]

        bo_inside = []
        bo_type_inside = []
        at_ea_inside = []
        
        # indi_list(self, list_of_indi, indi_type, desc)

        for box_office in box_office_list.list_of_indi:
            if block_current.inside_block(box_office):
                bo_inside.append(box_office)
        bo_inside_list = indi_list(bo_inside, box_office_list.indi_type,desc)        
        
        for bo_type in bo_type_list.list_of_indi:
            if block_current.inside_block(bo_type):
                bo_type_inside.append(bo_type)  
        bo_type_inside_list = indi_list(bo_type_inside, bo_type_list.indi_type,desc)        

        for at_ea in at_ea_list.list_of_indi:
            if block_current.inside_block(at_ea):
                at_ea_inside.append(at_ea)   
        at_ea_inside_list = indi_list(at_ea_inside, at_ea_list.indi_type,desc)
        print(at_ea_inside)
                
        # check if there is any "@ ea of" inside the block
        if at_ea_inside_list.length > 0 and bo_type_inside_list.length > 0: 
            print("at ea appears !!!")

            if len(bo_type_inside_list.set) == 2 and bo_type_inside_list.length%2 == 0:

                print("this is 2 bo type!!!")

            # 2 Things need to do here:
            # 1. repeat "WWBO" and "DBO"
            # 2. change relation into "OR" and "AND" combination.

                bonus_list_all = bonus_list_all + np.repeat(bonus.indi,bo_inside_list.length).tolist()
                
                
                for i in range(len(bo_inside)/2):
                    bo_type_list_all = bo_type_list_all + [bo_type_inside_list.list[0],bo_type_inside_list.list[1]]
                    relation_string = relation_string+" AND ("+str(start_num+1+2*i)+\
                    " OR "+str(start_num+2+2*i) + ")"
 
                    
            #if len(set(bo_type_inside)) == 1:
            else:
                
                print("this is 1 bo type!!!")
                
                print(bo_inside_list.list)
                bonus_list_all = bonus_list_all + np.repeat(bonus.indi,bo_inside_list.length).tolist()
                
                bo_type_list_all = bo_type_list_all + \
                np.repeat(bo_type_inside_list.list[0],bo_inside_list.length).tolist()
                
                relation_string = relation_string + "("+\
                ") AND (".join([str(i) for i in range(start_num+1, start_num+bo_inside_list.length+1)]) + ")"


     
                
        else:
            # no "@ ea of"
            ############################
            # Attention!!! If there is no "@ ea of" but are multiple box office amounts, 
            # I choose to leave the whole bonus column blank, if you want to change,
            # add (bo_inside_list.length - 1) times "0" maybe a good chioce.
            ############################
            bonus_list_all = bonus_list_all + [bonus.indi]
            bo_type_list_all = bo_type_list_all + bo_type_inside_list.list
            relation_string = relation_string + " AND (" + str(start_num+1) + ")"

    print("Bonus types " + str(len(bo_type_list_all)))
    print("Bonus types " + str(bo_type_list_all))


    # see if there is correct to split
    print("length of bo_type_list_all + "  + str(len(bo_type_list_all)))
    if len(bo_type_list_all) == num:
        df.at[ind,"Compensation.Box_Office_Type"] = bo_type_list_all
        df.at[ind,"Compensation.Bonus_Amount"] = bonus_list_all
        df.at[ind,"Compensation.Box_Office_Relationship"] = relation_string.lstrip(" AND ")
        df.at[ind,"Compensation.Parsed_sentence"] =  [bo_money_back_transfer(m) + '  ' + bonus_money_back_transfer(n) \
                                                  for m,n in zip(box_office_list.list,bonus_list_all)]    

    return df

#The main parsing spliter
    
# for not Director or Writer
def bo_bonus_type_spliter(df, func, comp_id):
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index

    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        desc = row["Right_COMPENSATION_DESC"]        

        # Catch box office amount, bonus amount, box office type, "or", "@"

        box_office_list = box_office_regex(desc)
        bonus_list = bonus_regex(desc)
        bo_type_list = bo_type_regex(desc)
        or_list = or_regex(desc)
        at_ea_list = at_ea_of_regex(desc)
        
        #ea_lsit?
        #(Right_COMPENSATION_AMOUNT) @ ea of DBO = $50M, $60M, $70M, $80M, $90M, $100M, $110M, $120M, $130M, $140M
        #$250K @ WWBO = $240M, $300K @ WWBO = $225M, 
        #$350K @ WWBO = $270M, $400K @ WWBO = $285M, $450K @ WWBO = $300M, $500K @ ea of WWBO = $315M & $330M;

        ##############################################################
        #For type like  DBO                      Bonus
        #              150M                   $100,000

        if box_office_list.length > 1 and box_office_list.length == bonus_list.length and \
        len(bo_type_list.set)==1:
            df = type_spliter_1(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list)
            continue


        ##############################################################        
        #For type like DBO     WWBO*  Bonus 
        #              150M or 375M $200,000,
        #find how many times "or" have shown.

        if box_office_list.length > 1 and bonus_list.length >= 1 and \
        len(bo_type_list.set) == 2 and or_list.length >= 1 and \
        (or_list.length + bonus_list.length) == box_office_list.length:
            df = type_spliter_2(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list)
            continue

        ##############################################################
        # CUT BY (bouns amount)
        # !!!!!Test if "@ ea of" is at first position!!!!!!
        # if "@ ea of" is at first position, insert "Right_COMPENSATION_AMOUNT" at first position of desc
        # (Right_COMPENSATION_AMOUNT) @ ea of DBO = $50M, $60M, $70M, $80M, $90M, $100M, $110M, $120M, $130M, $140M
        # $250K @ WWBO = $240M, $300K @ WWBO = $225M, 
        # $350K @ WWBO = $270M, $400K @ WWBO = $285M, $450K @ WWBO = $300M, $500K @ ea of WWBO = $315M & $330M;
        # (bouns amount) @ ea of (BO Type) = (BO1), (BO2), (BO3)

        if at_ea_list.length > 0 and box_office_list.length > 1 and bonus_list.length >= 1:           
            df = type_spliter_3(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list)
            continue

    return df.reset_index(drop=True)

def sole_shared_spliter(df, ind, func):
    
    rows = df.loc[ind]
    relation = rows["Compensation.Box_Office_Relationship"].values[0]
    bonus_amount = rows["Compensation.Bonus_Amount"].values
    bo_index = rows["Compensation.Box_Office_Index"].values
 
    num = rows.shape[0]
    
    # 2 means double the rows
    df = line_spliter(df, ind, 2)
    

    new_relation = ""
    for i in relation:
        if i.isdigit():
            new_relation = new_relation + str(int(i) + num)
        else:
            new_relation = new_relation + i
    
    new_bonus_amount = bonus_amount/2
    
    new_bo_index = bo_index + num
    
    df.at[ind, "Compensation.Box_Office_Relationship"] = np.repeat(relation,num).tolist() + np.repeat(new_relation,num).tolist()
    df.at[ind, "Compensation.Bonus_Amount"] = bonus_amount.tolist() + new_bonus_amount.tolist()
    df.at[ind, "Compensation.Box_Office_Index"] = bo_index.tolist() + new_bo_index.tolist()
    
    
    if func == "Writer":
        df.at[ind, "Compensation.Bonus_Type"] = np.repeat("Writing Credit Bonus", num*2)
        df.at[ind, "Compensation.Sole_Shared"] = np.repeat("Sole Credit Bonus", num).tolist() +\
        np.repeat("Shared Credit Bonus", num).tolist()

    elif func == "Director":
        df.at[ind, "Compensation.Bonus_Type"] = np.repeat("Sole Directing Credit", num).tolist() + \
        np.repeat("Shared Directing Credit", num).tolist()
    
     
    return df
    
    
def sole_shared_only_two_spliter(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list):
    
    row = df.loc[ind]
    bonus_amount = bonus_list.list

    # 2 means double the rows
    df = line_spliter(df, ind, 2)

    df.at[ind, "Compensation.Bonus_Amount"] = bonus_amount
    
    if func == "Writer":
        df.at[ind, "Compensation.Bonus_Type"] = np.repeat("Writing Credit Bonus", 2)
        df.at[ind, "Compensation.Sole_Shared"] = ["Sole Credit Bonus","Shared Credit Bonus"]

    elif func == "Director":
        df.at[ind, "Compensation.Bonus_Type"] = ["Sole Directing Credit","Shared Directing Credit"]
    
    return df

# for only Director or Writer
def director_writer_type_spliter(df, func, comp_id):
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index

    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        desc = row["Right_COMPENSATION_DESC"] 
        
        # remove pay WGA
        desc = re.sub(r'(?i)pay.{,10}WGA', '', desc)
       
        amount = row["Right_COMPENSATION_AMOUNT"]

        # Catch box office amount, bonus amount, box office type, "or", "@"

        box_office_list = box_office_regex(desc)
        bonus_list = bonus_regex(desc)
        bo_type_list = bo_type_regex(desc)
        or_list = or_regex(desc)
        at_ea_list = at_ea_of_regex(desc)
        sole_shared_flag = share_sole_credit_regex(desc)
        share_list = share_credit_regex(desc)
        sole_list = sole_credit_regex(desc)
        
        
        ################################################################
        # 
        if share_list.length > 0 and sole_list.length > 0 and bonus_list.length == 2:
            df = sole_shared_only_two_spliter(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list)
            continue
        
         # (Right_COMPENSATION_AMOUNT) flat for sole writing credit or $112,500 flat for shared credit
        if share_list.length > 0 and sole_list.length > 0 and bonus_list.length == 1 and amount > 0:
            bonus_list.add_at_left(indicator(amount, "bonus", -2, -1, desc))
            df = sole_shared_only_two_spliter(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list)
            continue

        ##############################################################
        #For type like  DBO                      Bonus
        #              150M                   $100,000

        if box_office_list.length > 1 and box_office_list.length == bonus_list.length and \
        len(bo_type_list.set)==1:
            df = type_spliter_1(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list)
            if sole_shared_flag:
                df = sole_shared_spliter(df, ind, func)
            continue

        ##############################################################        
        #For type like DBO     WWBO*  Bonus 
        #              150M or 375M $200,000,
        #find how many times "or" have shown.

        if box_office_list.length > 1 and bonus_list.length >= 1 and \
        len(bo_type_list.set) == 2 and or_list.length >= 1 and \
        (or_list.length + bonus_list.length) == box_office_list.length:
            df = type_spliter_2(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list)
            if sole_shared_flag:
                df = sole_shared_spliter(df, ind, func)
            continue

        ##############################################################
        # CUT BY (bouns amount)
        # !!!!!Test if "@ ea of" is at first position!!!!!!
        # if "@ ea of" is at first position, insert "Right_COMPENSATION_AMOUNT" at first position of desc
        # (Right_COMPENSATION_AMOUNT) @ ea of DBO = $50M, $60M, $70M, $80M, $90M, $100M, $110M, $120M, $130M, $140M
        # $250K @ WWBO = $240M, $300K @ WWBO = $225M, 
        # $350K @ WWBO = $270M, $400K @ WWBO = $285M, $450K @ WWBO = $300M, $500K @ ea of WWBO = $315M & $330M;
        # (bouns amount) @ ea of (BO Type) = (BO1), (BO2), (BO3)

        if at_ea_list.length > 0 and box_office_list.length > 1 and bonus_list.length >= 1:           
            df = type_spliter_3(df, func, ind, box_office_list, bonus_list, bo_type_list, or_list, at_ea_list)
            if sole_shared_flag:
                df = sole_shared_spliter(df, ind, func)
            continue

    return df.reset_index(drop=True)

#Main Iteration Function
def split_pasting_func(func_df,func):
    df = func_df.copy()
    id_list = df.loc[df["Index"] == "New"]["Right_COMPENSATION_ID"].unique()
    for comp_id in id_list:
        print(func)
        print(comp_id)
        
        if func ==  "Writer":
            df = compensation_type_paster(df, comp_id)
            df = writer_bonus_type_paster(df, comp_id)
            df, flag = writer_pp_np_num_spliter(df,comp_id)
            if not flag:
                df = director_writer_type_spliter(df,func,comp_id)
                
        elif func == "Director":
            df = compensation_type_paster(df, comp_id)
            df,flag = pp_np_num_paster(df, comp_id)
            if not flag:
                df = director_writer_type_spliter(df,func,comp_id)           
            
        else:
            df = compensation_type_paster(df, comp_id)
            df,flag = pp_np_num_paster(df, comp_id)
            if not flag:
                df = bo_bonus_type_spliter(df, func,comp_id)

        
    return df

#Using this dictionary to reorder the output columns.
column_order_list = {'Actor':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Bonus_Type","Compensation.Royalty_%",
                              "Compensation.PP_%","Compensation.PP_np/gp","Compensation.Deferment_Amount",
                              "Compensation.Oscar_Bonus_Amount","Compensation.Golden_Globe_Bonus_Amount",
                              "Compensation.On_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
                              "Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount",
                              "Compensation.Under_Budget_Bonus_ProRata","Compensation.Over_Budget_Penalty_%",
                              "Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
                              "Compensation.Box_Office_Relationship","Compensation.Box_Office_Type",
                              "Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                              "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                              "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                              "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                              "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                              "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                     'Casting':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Bonus_Type","Compensation.Royalty_%",
                              "Compensation.PP_%","Compensation.PP_np/gp","Compensation.Deferment_Amount",
                              "Compensation.Oscar_Bonus_Amount","Compensation.Golden_Globe_Bonus_Amount",
                              "Compensation.On_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
                              "Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount",
                              "Compensation.Under_Budget_Bonus_ProRata","Compensation.Over_Budget_Penalty_%",
                              "Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
                              "Compensation.Box_Office_Relationship","Compensation.Box_Office_Type",
                              "Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                              "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                              "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                              "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                              "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                              "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                     'Consultant':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID",
                                   "COMPENSATION_AMOUNT","COMPENSATION_DESC","COMPENSATION_TYPE",
                                   "Compensation.Bonus_Type","Compensation.Royalty_%","Compensation.PP_%",
                                   "Compensation.PP_np/gp","Compensation.Prod_Bonus_Amount",
                                   "Compensation.Deferment_Amount","Compensation.Oscar_Bonus_Amount",
                                   "Compensation.Golden_Globe_Bonus_Amount","Compensation.On_Budget_Bonus_Amount",
                                   "Compensation.Under_Budget_Bonus_Direct_Cost","Compensation.Under_Budget_Bonus_Qualifier",
                                   "Compensation.Under_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_ProRata",
                                   "Compensation.Over_Budget_Penalty_%","Compensation.Over_Budget_Amount",
                                   "Compensation.Over_Budget_ProRata","Compensation.Box_Office_Relationship",
                                   "Compensation.Box_Office_Type","Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                                   "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                                   "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                                   "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                                   "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                                   "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                     'Director':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Bonus_Type","Compensation.Royalty_%",
                              "Compensation.PP_%","Compensation.PP_np/gp","Compensation.Deferment_Amount",
                              "Compensation.Oscar_Bonus_Amount","Compensation.Golden_Globe_Bonus_Amount",
                              "Compensation.On_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
                              "Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount",
                              "Compensation.Under_Budget_Bonus_ProRata","Compensation.Over_Budget_Penalty_%",
                              "Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
                              "Compensation.Box_Office_Relationship","Compensation.Box_Office_Type",
                              "Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                              "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                              "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                              "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                              "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                              "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                     'Financier':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Bonus_Type","Compensation.Royalty_%",
                              "Compensation.PP_%","Compensation.PP_np/gp","Compensation.Deferment_Amount",
                              "Compensation.Oscar_Bonus_Amount","Compensation.Golden_Globe_Bonus_Amount",
                              "Compensation.On_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
                              "Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount",
                              "Compensation.Under_Budget_Bonus_ProRata","Compensation.Over_Budget_Penalty_%",
                              "Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
                              "Compensation.Box_Office_Relationship","Compensation.Box_Office_Type",
                              "Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                              "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                              "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                              "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                              "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                              "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                     'Termdeal':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID",
                                 "COMPENSATION_AMOUNT","COMPENSATION_DESC","COMPENSATION_TYPE",
                                 "Compensation.Bonus_Type","Compensation.Royalty_%","Compensation.PP_%",
                                 "Compensation.PP_np/gp","Compensation.Deferment_Amount","Compensation.Oscar_Bonus_Amount",
                                 "Compensation.Golden_Globe_Bonus_Amount","Compensation.On_Budget_Bonus_Amount",
                                 "Compensation.Under_Budget_Bonus_Direct_Cost","Compensation.Under_Budget_Bonus_Qualifier",
                                 "Compensation.Under_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_ProRata",
                                 "Compensation.Over_Budget_Penalty_%","Compensation.Over_Budget_Amount",
                                 "Compensation.Over_Budget_ProRata","Compensation.Box_Office_Relationship",
                                 "Compensation.Box_Office_Type","Compensation.Parsed_sentence",
                                 "Compensation.Bonus_Amount","Compensation.Box_Office_Qualifier",
                                 "Compensation.Box_Office_Amount","Compensation.Box_Office_Qualifier2",
                                 "Compensation.Box_Office_Index","Compensation.Bonus_Compensation_Type",
                                 "TERM_DEAL_DURATION","#_Duration","Qualifier_Duration","Index",
                                 "Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                                 "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID",
                                 "Right_FUNCTION","Right_PROJECT_ID"],
                     'RightsIn':["DARTS_DIVISION","COMPENSATION_ID","COMPENSATION_AMOUNT",
                                 "COMPENSATION_DESC","COMPENSATION_TYPE","DEAL_ID","FUNCTION","PROJECT_ID",
                                 "Compensation.Bonus_Type","Compensation.Royalty_%","Compensation.PP_%",
                                 "Compensation.PP_np/gp","Compensation.Deferment_Amount",
                                 "Compensation.Oscar_Bonus_Amount","Compensation.Golden_Globe_Bonus_Amount",
                                 "Compensation.On_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
                                 "Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount",
                                 "Compensation.Under_Budget_Bonus_ProRata","Compensation.Over_Budget_Penalty_%",
                                 "Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
                                 "Compensation.Box_Office_Relationship","Compensation.Box_Office_Type",
                                 "Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                                 "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                                 "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                                 "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                                 "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                                 "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                     'RightsOut':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Bonus_Type","Compensation.Royalty_%",
                              "Compensation.PP_%","Compensation.PP_np/gp","Compensation.Deferment_Amount",
                              "Compensation.Oscar_Bonus_Amount","Compensation.Golden_Globe_Bonus_Amount",
                              "Compensation.On_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
                              "Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount",
                              "Compensation.Under_Budget_Bonus_ProRata","Compensation.Over_Budget_Penalty_%",
                              "Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
                              "Compensation.Box_Office_Relationship","Compensation.Box_Office_Type",
                              "Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                              "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                              "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                              "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                              "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                              "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                    'Generic_Functions':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID",
                                         "COMPENSATION_AMOUNT","COMPENSATION_DESC","COMPENSATION_TYPE",
                                         "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID",
                                         "Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                                         "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                     'Producer':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID",
                                 "COMPENSATION_AMOUNT","COMPENSATION_DESC","COMPENSATION_TYPE",
                                 "Compensation.Bonus_Type","Compensation.Royalty_%","Compensation.PP_%",
                                 "Compensation.PP_np/gp","Compensation.Deferment_Amount","Compensation.Oscar_Bonus_Amount",
                                 "Compensation.Golden_Globe_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
                                 "Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount",
                                 "Compensation.Under_Budget_Bonus_ProRata","Compensation.Over_Budget_Penalty_%",
                                 "Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
                                 "Compensation.Box_Office_Relationship","Compensation.Box_Office_Type",
                                 "Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                                 "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                                 "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                                 "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                                 "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                                 "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID",
                                 "Right_FUNCTION","Right_PROJECT_ID"],
                     'Writer':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID",
                               "COMPENSATION_AMOUNT","COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Bonus_Type",
                               "Compensation.Royalty_%","Compensation.Sole_Shared","Compensation.Writing_Credit_np/gp",
                               "Compensation.Writing_Credit_%","Compensation.Deferment_Amount",
                               "Compensation.Oscar_Bonus_Amount","Compensation.Golden_Globe_Bonus_Amount",
                               "Compensation.On_Budget_Bonus_Amount","Compensation.Under_Budget_Bonus_Direct_Cost",
                               "Compensation.Under_Budget_Bonus_Qualifier","Compensation.Under_Budget_Bonus_Amount",
                               "Compensation.Under_Budget_Bonus_ProRata","Compensation.Over_Budget_Penalty_%",
                               "Compensation.Over_Budget_Amount","Compensation.Over_Budget_ProRata",
                               "Compensation.Box_Office_Relationship","Compensation.Box_Office_Type",
                               "Compensation.Parsed_sentence","Compensation.Bonus_Amount",
                               "Compensation.Box_Office_Qualifier","Compensation.Box_Office_Amount",
                               "Compensation.Box_Office_Qualifier2","Compensation.Box_Office_Index",
                               "Compensation.Bonus_Compensation_Type","Index","Right_DARTS_DIVISION",
                               "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT","Right_COMPENSATION_DESC",
                               "Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"]}    


# Main loop
for func in dict_combined.keys():
    dict_combined[func] = split_pasting_func(dict_combined[func],func)
    #One more step, choose excat columns with exact order.
    dict_combined[func] = dict_combined[func][column_order_list[func]]
    
#Writing excel
for func in dict_combined.keys():
    with pd.ExcelWriter(output_position + '/' + date + "_Contingent_" + func + "_" +
                              "Parsed.xlsx", engine='xlsxwriter') as writer:
        dict_combined[func].to_excel(writer, sheet_name='Contingent_Compensation',index = False)
        
print("Parsing complete.")
    
    
