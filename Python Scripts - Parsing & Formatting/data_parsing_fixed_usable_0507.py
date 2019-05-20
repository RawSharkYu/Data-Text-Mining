# -*- coding: utf-8 -*-
"""
Created on Mon Apr  8 15:18:03 2019

@author: AYu5
"""

import sys
import pandas as pd
import numpy as np
import re
from decimal import Decimal
import datetime as dt


# Change following 3 variables into current positions and date

input_position = (r'C:\Users\ayu5\Desktop\Old_Data_Parsing\data_parsing_0128\data_parsing_raw\Use Every Time_Output Data - Copy')
output_position = (r'C:\Users\ayu5\Desktop\PythonForDP\output')
date =  '05072019'


# Don't change any the codes below unless necessary and certain

input_position = input_position.replace('\\','/')
output_position = output_position.replace('\\','/')

function_name_list = ['Actor','Casting','Consultant','Director','Financier',
                 'Rights','Generic_Functions','Producer','Researcher','Writer']

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

#First step, read files
dict_combined = {}
for func in function_name_list:
    if func == 'Rights':
        print(func)
        existed_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_Rights_Fixed_In&Out_Existed.xlsx")
        new_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_Rights_Fixed_In&Out_New.xlsx")
        existed_df = pd.read_excel(existed_file,"Sheet1")
        new_df = pd.read_excel(new_file,"Sheet1")
    else:
        existed_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_" + func + "_Fixed_Existed.xlsx")
        new_file = pd.ExcelFile(input_position + '/' + "Delta_Compensation_" + func + "_Fixed_New.xlsx")
        existed_df = pd.read_excel(existed_file,"Sheet1")
        new_df = pd.read_excel(new_file,"Sheet1")
    
    combined = pd.concat([existed_change(existed_df),new_change(new_df)],sort=False,ignore_index=True)
    dict_combined[func] = combined
    
    
    
    
    
#Variable Preparation
    
#Add one temporary column named "split_desc", which is used to specify split content for each line, which will be droped at the end.

#Date time type convert. For "Compensation.Start_Date","Rights Start Date",
#"Compensation Commitment Date","Start_Date","Expiry_Date","Purchase_Date","Reversion_Date", "Date Paid".

date_columns = ["Compensation.Start_Date","Rights Start Date",
                "Compensation Commitment Date","Start_Date","Expiry_Date",
                "Purchase_Date","Reversion_Date", "Date Paid"]
 
print("Transforming date type...")

for key,df in dict_combined.items():
    
    print(key)
    df['split_desc'] = df["Right_COMPENSATION_DESC"]
    
    for col in date_columns:
        print(col)
        if col in df.columns:
            df[col] = pd.to_datetime(df[col],errors='coerce')
    
    #df["Right_COMPENSATION_ID"] = df["Right_COMPENSATION_ID"].fillna(-1)
    #df["Right_COMPENSATION_ID"] = df["Right_COMPENSATION_ID"].astype('int64')
    #df["Right_COMPENSATION_ID"] = df["Right_COMPENSATION_ID"].astype('str')
    #df["Right_COMPENSATION_ID"] = df["Right_COMPENSATION_ID"].replace("-1", np.nan)
    
    #df['Compensation.Start_Date'] = df['Compensation.Start_Date'].astype('str')
    
print("Date type transforming succesfully done.")    
    
#General Function Preparation
    
    
# 1.line spliter
def line_spliter(orig_df,index,num):
    df = orig_df.copy()
    row = df.loc[index]
    print(num - 1)
    for j in range(num - 1):
        df = df.append(row)
    return df

# Transfer money string format from "$1.2K" into "1200"
def money_transfer(money_str):
    string  = money_str
    num = Decimal(re.sub(r'[^\d.]', '', string))
    if 'k' in string or 'K' in string:
        num = num * 1000
    if 'm' in string or 'M' in string:
        num = num * 1000000
    return num




#Data Parsing
    
#1.First step, copy and paste "Right_COMPENSATION_TYPE".
def comp_type_paster(df, comp_id):
    row = df[df["Right_COMPENSATION_ID"] == comp_id]
    ind = row.index[0]
    df.at[ind, "Compensation.Compensation_Type"] = df.loc[ind, 'Right_COMPENSATION_TYPE']
    return df

#2.second step, "Orig deal"
# Remove string behind "Orig deal"
# Not for "New", just for "Updated" at now.
# Not used now.
def orig_deal_paster(df, comp_id):
    row = df[df["Right_COMPENSATION_ID"] == comp_id]
    ind = row.index[0]
    desc = row['split_desc'].values[0]
    # To see if there is "orig deal"
    sub_strs = re.findall(r"(?i)(.*)(Origi?n?a?l? deal.*)", desc)
    if len(sub_strs) == 1:
        df.at[ind, 'split_desc'] = sub_strs[0][0]
    return df


#3.third step, "overage".
def overage_spliter(df, comp_id):
    # For now, only 1 row
    row = df[df["Right_COMPENSATION_ID"] == comp_id]
    index = row.index[0]
    desc = row['split_desc'].values[0]
    # To see if there is "orig deal"
    sub_strs = re.findall(r"(?i)(.*)(overage.{4,})", desc)
    #split line
    if len(sub_strs) == 1:
        num = 2
        df = line_spliter(df, index, num)
        df.at[index,'split_desc'] = sub_strs[0]
    return df.reset_index(drop=True)


#4.fourth step, "Payment_Type"
def Payment_Type_Regex(str_test):    
    
    # mode1:(25/25/25/25)
    sub_strs = re.findall(r"(?i)25\%(.*)25\%(.*)25\%(.*)25\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"25/25/25/25 Split",['0.25','0.25','0.25','0.25'],sub_strs)
    
    # mode2:(50/50)
    sub_strs = re.findall(r"(?i)50\%(.*)50\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"Equal Installments",['0.5','0.5'],sub_strs)
    
    sub_strs = re.findall(r"(?i)1\/2(.*)1\/2(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"Equal Installments",['0.5','0.5'],sub_strs)
    
    # mode3:(20/60/10/10)
    sub_strs = re.findall(r"(?i)20\%(.*)60\%(.*)10\%(.*)10\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"20/60/10/10 Split",['0.2','0.6','0.1','0.1'],sub_strs)

    # mode4:(20/60/20)
    sub_strs = re.findall(r"(?i)20\%(.*)60\%(.*)20\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"20/60/20 Split",['0.2','0.6','0.2'],sub_strs)
    
    # mode5:(80/10/10)
    sub_strs = re.findall(r"(?i)80\%(.*)10\%(.*)10\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"80/10/10 Split",['0.8','0.1','0.1'],sub_strs)  

    # mode6:(15/60/12.5/12.5)
    sub_strs = re.findall(r"(?i)15\%(.*)60\%(.*)12.5\%(.*)12.5\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"15/60/12.5/12.5 Split",['0.15','0.6','0.125','0.125'],sub_strs)  
    
    # mode7:(60/10/10/10/10)
    sub_strs = re.findall(r"(?i)60\%(.*)10\%(.*)10\%(.*)10\%(.*)10\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"60/10/10/10/10 Split",['0.6','0.1','0.1','0.1','0.1'],sub_strs) 
    
    
    # mode8:(1/3 1/3 1/3)
    sub_strs = re.findall(r"(?i)1\/3(.*)1\/3(.*)1\/3(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"1/3 1/3 1/3 split",['1/3','1/3','1/3'],sub_strs) 
    
    # no mode caught
    return (False,"",[],[])



def Payment_Type_spliter(df, comp_id):
    # For now, there may be more than 1 row
    # do it for every row
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index

    if len(indexs) > 0:
        for ind in indexs:
            print(comp_id)
            
            row = df.loc[ind]
            
            desc = row["split_desc"]
            flag, type_name, percen_total, sub_strs = Payment_Type_Regex(desc)
            
            if flag:

                df = line_spliter(df, ind, len(percen_total))
                df.at[ind, "Compensation.Payment_Type"] = type_name#np.repeat(type_name,len(percen_total))
                df.at[ind, "Compensation.%_Total"] = percen_total
                df.at[ind, "split_desc"] = sub_strs[0]
        
    return df.reset_index(drop=True)


#5.Fifth step, other non-split-line columns to all functions.
def start_quali_condi_regex(str_test):
    
    print(str_test)
    
    # mode1:"Over Scheduled Principal Photography"
    
    # pybl over sched, payable over schedule, over sched period of prin photo
    sub_strs = re.findall(r"(?i)over\sschedu?l?e?", str_test)
    if len(sub_strs) > 0:
        return (True,"Over","Scheduled Principal Photography")
    
    # over princ photog, prin photo, over prin photo
    sub_strs = re.findall(r"(?i)princ?i?p?a?l?\sphoto", str_test)
    if len(sub_strs) > 0:
        return (True,"Over","Scheduled Principal Photography")
    
    # sched of photo comm
    sub_strs = re.findall(r"(?i)schedu?l?e?\sof\sphoto", str_test)
    if len(sub_strs) > 0:
        return (True,"Over","Scheduled Principal Photography")    
    
    
    # model2:"Over Pre-Production"
    # "over 8 wks pre-prod, over 8 wks preprod, over preproduction"
    sub_strs = re.findall(r"(?i)over\s.{,12}pre\-?produ?c?t?i?o?n?", str_test)
    if len(sub_strs) > 0:
        return (True,"Over","Pre-Production")
    
    
    
    # model3: "Over Production"
    # 'over production'
    sub_strs = re.findall(r"(?i)over\s.{,12}produ?c?t?i?o?n?", str_test)
    if len(sub_strs) > 0:
        return (True,"Over","Production")    
    
    
    
    # model4:"On Completion of Delivery"
    # "on compl del"
    sub_strs = re.findall(r"(?i)on\scomple?t?i?o?n?e?\sdeli?v?e?r?y?", str_test)
    if len(sub_strs) > 0:
        return (True,"On","Completion of Delivery") 
    
    
    # model5:"On Completion of Dub/Score"
    # "on compl dub/score"
    sub_strs = re.findall(r"(?i)on\scomple?t?i?o?n?e?\sdub\/score", str_test)
    if len(sub_strs) > 0:
        return (True,"On","Completion of Dub/Score")    
    
    
    # model6:"On Election to Proceed/Abandon"
    # "on election to proceed/abandon, elects proceed/ abandon, upon election to proceed or abandonment"
    sub_strs = re.findall(r"(?i)proceed\W{,2}o?r?\W{,2}abandon", str_test)
    if len(sub_strs) > 0:
        return (True,"On","Election to Proceed/Abandon")    
    
    
    # model7:"On Commencement of Services"
    # "on commencement of screenwriter's services, on commencement of opt. services,
    # on commencement of writer's services, comm of Teitler/Weber/Field's svcs"
    sub_strs = re.findall(r"(?i)comme?n?c?e?m?e?n?t? of .{,15} se?r?vi?ce?s", str_test)
    if len(sub_strs) > 0:
        return (True,"On","Commencement of Services")  
    
    
    # model8:"On Delivery of Answer Print"
    # "answer print"
    sub_strs = re.findall(r"(?i)answer print", str_test)
    if len(sub_strs) > 0:
        return (True,"On","Delivery of Answer Print")      
    
    
    # model 9:"On Satisfaction of Conditions Precedent"
    # "on satisfaction of conditions precedent,satis. of cond. prec.,
    # satisfaction of conditions precedent, satisfn of cond precedent,"
    sub_strs = re.findall(r"(?i)satisf?a?c?t?i?o?n?\.? of condi?t?i?o?n?s?\.? prece?d?e?n?t?\.?", str_test)
    if len(sub_strs) > 0:
        return (True,"On","Delivery of Answer Print")    
    
    
    # model10:"On Execution of Agreement"
    # "on execution of long form agreement, upon execution of the agreement, on execution, upon exec, on exec"
    sub_strs = re.findall(r"(?i)on execu?t?i?o?n?", str_test)
    if len(sub_strs) > 0:
        return (True,"On","Execution of Agreement")     
    
    # model11:"Later of"
    # "later of"
    sub_strs = re.findall(r"(?i)later of", str_test)
    if len(sub_strs) > 0:
        return (True,"Later of","")
    
    # model12:"Earlier of"
    # "earlier of"
    sub_strs = re.findall(r"(?i)earlier of", str_test)
    if len(sub_strs) > 0:
        return (True,"Earlier of","")    
    
    
    
    # no mode caught
    return (False,"","")


# For "Compensation.Start_Qualifier" and 
# "Compensation.Payment_Start_Condition" or "Compensation.Service_Start_Condition" (Director)
def start_quali_condi_paster(df, comp_id):
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    if len(indexs) > 0:
        for ind in indexs:
            print(comp_id)
            row = df.loc[ind]
            desc = row["split_desc"]
            print(desc)
            flag, start_quali, start_condi = start_quali_condi_regex(desc)
            
            if flag:
                df.at[ind, "Compensation.Start_Qualifier"] = start_quali
                if 'Compensation.Payment_Start_Condition' in df.columns:
                    df.at[ind, 'Compensation.Payment_Start_Condition'] = start_condi
                elif "Compensation.Service_Start_Condition" in df.columns:
                    df.at[ind,"Compensation.Service_Start_Condition"] = start_condi
    
    return df


#For Compensation.%_Amount
# for Compensation.%_Amount
def per_amount_spliter(df, comp_id):
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    if len(indexs) > 0:
        for ind in indexs:
            print(comp_id)
            row = df.loc[ind]
            desc = row["split_desc"]
            
            result = re.findall(r"(?i)\$\W?([1-9]\d{0,2}[\,\.]?\d{,3}K?M?)",desc)
            
            #if len(result) == 1:
                
            #    df.at[ind,"Compensation.%_Amount"] = money_transfer(result[0])
            
            if len(result) > 1:
                pos = [0]
                moneys = []
                for m in re.finditer(r"(?i)\$\W?([1-9]\d{0,2}[\,\.]?\d{,3}K?M?)",desc):
                    pos.append(m.start())
                    pos.append(m.end())
                    moneys.append(money_transfer(m.groups()[0]))
                    
                pos.append(len(desc))
                i = 0
                flags = 0
                split_desc_list = []
                while i+1 <= len(pos):
                    start_point = pos[i]
                    print(start_point)
                    end_point = pos[i+1]
                    print(end_point)
                    print(desc[start_point:end_point])
                    flag, start_quali, start_condi = start_quali_condi_regex(desc[start_point:end_point])
                    if flag:
                        split_desc_list.append(desc[start_point:end_point])
                        flags = flags + 1
                    i = i + 2

                
                if len(moneys) == flags:
                    print("cut_amount")
                    print(split_desc_list)
                    df = line_spliter(df, ind, flags)
                    df.at[ind,'Compensation.%_Amount'] = moneys
                    df.at[ind,"split_desc"] = split_desc_list
            else:
                return df.reset_index(drop=True)
    
    return df.reset_index(drop=True)   


# For 'Compensation.Duration_#' & 'Compensation.Duration_Freq'.
def dura_num_freq_paster(df, comp_id):
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    
    if len(indexs) > 0:
        for ind in indexs:
            print(comp_id)
            
            row = df.loc[ind]
            
            desc = row["split_desc"]
            capture = 'session|week|wk|month|mo|year|yr|day'
            result = re.findall(r"(?i)\s(?=(\d{1,2}\.?\d?)\W("+capture+r"))\b",desc)
            if len(result) > 0:
                df.at[ind, 'Compensation.Duration_#'] = result[0][0].replace(",","")
                df.at[ind, 'Compensation.Duration_Freq'] = result[0][1].replace("day","Day")\
                .replace("session","Session").replace("wk","Week").replace("week","Week")\
                .replace("mo","Month").replace("month","Month").replace("Monthnth","Month")\
                .replace("yr","Year").replace("year","Year")
            else:
                continue; 
    
    return df


# For 'Compensation.Rate' & 'Compensation.Rate_Freq'.
def rate_num_freq_paster(df, comp_id):
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    
    if len(indexs) > 0:
        for ind in indexs:
            print(comp_id)
            
            row = df.loc[ind]
            
            desc = row["split_desc"]
            capture = 'session|week|wk|month|mo|year|yr|day'
            result = re.findall(r"(?i)\$\W?([1-9]\d{0,2}[\,\.]?\d{,3}K?)\W?(per|\/)\W?("+capture+r")",desc)
            
            
            # remove caught "$" + number, to help "%_Amount" catching
            for m in re.finditer(r"(?i)\$\W?([1-9]\d{0,2}[\,\.]?\d{,3}K?)\W?(per|\/)\W?("+capture+r")",desc):
                desc = desc[:m.start()]+desc[m.end():]
            
            df.at[ind, 'split_desc.Rate'] = desc
            
            
            
            if len(result) > 0:
                
                value = result[0][0]
                
                num = money_transfer(value)
                
                df.at[ind, 'Compensation.Rate'] = num
                df.at[ind, 'Compensation.Rate_Freq'] = result[0][2].replace("day","Day")\
                .replace("session","Session").replace("wk","Week").replace("week","Week")\
                .replace("mo","Month").replace("month","Month").replace("Monthnth","Month")\
                .replace("yr","Year").replace("year","Year")
            else:
                continue; 
                
    
    
    return df

# For 'Compensation.Start_Date'
# from/commencing/comm \b XX/XX/XX
# Not for two date time
def start_date_paster(df, comp_id):
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    
    if len(indexs) > 0:
        for ind in indexs:
            print(comp_id)
            row = df.loc[ind]
            desc = row["split_desc"]
            result = re.findall(r"(?i)(START DATE|comm|commencing|TERM)\W{,3}(\d{1,2}\/\d{1,2}\/\d{2,4})",desc)
            if len(result) > 0:
                df.at[ind, 'Compensation.Start_Date'] = result[0][1]
            else:
                continue; 
    
    return df


# 6.Sixth step, other non-split-line columns related each specific functions.
def fee_type_regex(func_name,desc):
    if func_name == "Actor":
        # Acting fee or Overage fee
        sub_strs = re.findall(r"(?i)overage", desc)
        if len(sub_strs) > 0:
            return "Overage Fee"
        else:
            return "Acting Fee"

            
    if func_name == "Casting":
        # Casting fee or Overage fee
        sub_strs = re.findall(r"(?i)overage", desc)
        if len(sub_strs) > 0:
            return "Overage Fee"
        else:
            return "Casting Fee"


    if func_name == "Consultant":
        # Consulting Fee or Overage fee
        sub_strs = re.findall(r"(?i)overage", desc)
        if len(sub_strs) > 0:
            return "Overage Fee"
        else:
            return "Consulting Fee"
    
            
    if func_name == "Director":
        # Directing Fee or Development Fee
        sub_strs = re.findall(r"(?i)development (fee|service)", desc)
        if len(sub_strs) > 0:
            return "Development Fee"
        else:
            return "Directing Fee"
    
            
    if func_name == "Financier":
        return np.nan
            
            
    if func_name == "GenericFunction":
        return np.nan
    
    if func_name == "Producer":
        # Producing Fee or Executive Producer Fee or Development Fee or Co-Producer Fee
        if len(re.findall(r"(?i)development (fee|service)", desc)) > 0:
            return "Development Fee"
        elif len(re.findall(r"(?i)Executive Prod", desc)) > 0 or len(re.findall(r"(?i)EP fee", desc)) > 0:
            return "Executive Producer Fee"
        elif len(re.findall(r"(?i)co\W?prod", desc)) > 0:
            return "Co-Producer Fee"            
        else:
            return "Producing Fee"
        
      
                
    if func_name == "Researcher":
        return np.nan
            
    if func_name == "Writer":
        return "Writing Fee"
      
    if func_name == "Rights":
        return np.nan


def fee_type_paster(df, func_name, comp_id):
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
        
    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        desc = row["split_desc"]
        
        # func_name: Actor,Casting,Consultant,Director,Financier,
        # GenericFunction,Producer,Researcher,Writer,Rights
        df.at[ind,"Compensation.Fee_Type"] = fee_type_regex(func_name,desc)          
            
    return df
    

#Compensation.Service_Type only for Producer.
def service_type_regex(desc):
    # Producer Services or Executive Producer Services or Co-Producer Services
    if len(re.findall(r"(?i)Executive Prod", desc)) > 0 or len(re.findall(r"(?i)EP fee", desc)) > 0:
        return "Executive Producer Services"
    elif len(re.findall(r"(?i)co\W?prod", desc)) > 0:
        return "Co-Producer Services"            
    else:
        return "Producer Services"    
    
def service_type_paster(df, func_name, comp_id):
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        desc = row["split_desc"]    
        df.at[ind,"Compensation.Service_Type"] = service_type_regex(desc)
    return df 

#7.Seventh step, unique columns (for Rights and Writer).
    
#Functions for Rights.
    
# for Compensatipon.Total_Guaranteed_Commitment
def total_guarant_paster(df, comp_id):
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index

    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        
        if row["Right_COMPENSATION_TYPE"] == "Guaranteed":
            df.at[ind,"Compensatipon.Total_Guaranteed_Commitment"] = row["Right_COMPENSATION_AMOUNT"]
        
    return df 

# for Compensation.Check_#
def check_number_paster(df, comp_id):
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        desc = row["split_desc"] 
        sub_strs = re.findall(r"(?i)ch?e?c?ks?\W{,3}(\d{3,})", desc)
        
        if len(sub_strs) > 0:
            df.at[ind,"Compensation.Check_#"] = sub_strs[0]
    return df 

# for Start_Date and Expiry_Date
# 1. XX/XX/XX thru/- XX/XX/XX
# 2. EXPIRES 3/29/10
def start_expiry_date_paster(df, comp_id):
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        desc = row["split_desc"] 
        
        #1.
        sub_strs = re.findall(r"(?i)(\d{1,2}\/\d{1,2}\/\d{1,2}) (\-|thru) (\d{1,2}\/\d{1,2}\/\d{1,2})", desc)
        
        if len(sub_strs) > 0:
            df.at[ind,"Start_Date"] = pd.to_datetime(sub_strs[0][0])
            df.at[ind,"Expiry_Date"] = pd.to_datetime(sub_strs[0][2])
            
        #2.
        sub_strs = re.findall(r"(?i)expire\w\W{,3}(\d{1,2}\/\d{1,2}\/\d{1,2})", desc)
        
        if len(sub_strs) > 0:
            df.at[ind,"Expiry_Date"] = pd.to_datetime(sub_strs[0])
            
    return df

# for Date Paid
# eg."**Paid via wire on 3/30/17.**"
def date_paid_paster(df, comp_id):
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        desc = row["split_desc"] 
        
        sub_strs = re.findall(r"(?i)\*\*.*(\d{1,2}\/\d{1,2}\/\d{1,2}).*\*\*", desc)
        
        if len(sub_strs) > 0:
            df.at[ind,"Date Paid"] = pd.to_datetime(sub_strs[0])
    
    
    return df

#Functions for Writer.
    
def writer_payment_type_regex(str_test):
    
    # mode1:(25/25/25/25)
    sub_strs = re.findall(r"(?i)25\%(.*)25\%(.*)25\%(.*)25\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"25/25/25/25 Split",['0.25','0.25','0.25','0.25'],sub_strs)

    # mode2:(50/50)
    sub_strs = re.findall(r"(?i)50\%(.*)50\%(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"Equal Installments",['0.5','0.5'],sub_strs)
    
    sub_strs = re.findall(r"(?i)1\/2(.*)1\/2(.*)", str_test)
    if len(sub_strs) > 0:
        return (True,"Equal Installments",['0.5','0.5'],sub_strs)
    
    return (False,"",[],[])

def Writer_Payment_Type_spliter(df, comp_id):
    # For now, there may be more than 1 row
    # do it for every row
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index

    if len(indexs) > 0:
        for ind in indexs:
            print(comp_id)
            
            row = df.loc[ind]
            
            desc = row["split_desc"]
            flag, type_name, percen_total, sub_strs = writer_payment_type_regex(desc)
            
            if flag:
                
                total_amount = df.loc[ind,"Right_COMPENSATION_AMOUNT"]

                df = line_spliter(df, ind, len(percen_total))
                df.at[ind, "Compensation.Payment_Type"] = type_name#np.repeat(type_name,len(percen_total))
                df.at[ind, "Compensation.%_Total"] = percen_total
                df.at[ind, "Compensation.%_Amount"] = np.repeat(total_amount * float(percen_total[0]),len(percen_total))
                df.at[ind, "split_desc"] = sub_strs[0]
        
    return df.reset_index(drop=True)

def writing_step_regex(str_test):
    
    catching_list = []
    
    order = '1st|first|2nd|second|3rd|third|4th|fourth|5th|fifth|6th|sixth'
    
    step_names = 'draft|polish|rewrite|revision|outline|treatment|R\/W'
    
    p = re.compile(r"(?i)("+order+r")?\s?("+step_names+r")")
    
    position_list = []
    
    for m in p.finditer(str_test):
        
        step = m.group()
        
        step = step.replace("draft","Draft").replace("polish","Polish")\
                             .replace('R/W','Rewrite').replace('rewrite','Rewrite')\
                             .replace("revision","Set of Revisions").replace("set of revisions","Set of Revisions")\
                             .replace("outline","Outline/Treatment").replace("treatment","Outline/Treatment")\
                            .replace("Outline","Outline/Treatment").replace("Treatment","Outline/Treatment")\
                            .replace("Outline/Outline/Treatment/Outline/Treatment","Outline/Treatment")\
                            .replace("first","1st").replace("second","2nd").replace("third",'3rd')\
                            .replace("First","1st").replace("Second","2nd").replace("Third",'3rd')\
                            .replace("fourth","4th").replace("fifth",'5th').replace('sixth','6th')\
                            .replace("Fourth","4th").replace("Fifth",'5th').replace('Sixth','6th')\
                            .lstrip(" ")
        
    
        catching_list.append(step)
        position_list.append(m.span())

    
    return (catching_list,position_list)

def writer_start_quali_condi_regex(str_test):
# For writer, start_quali_condi_regex is different, because there are only "On Commencement" and "On Delivery".

    catching_list = []
    
    p = re.compile(r"(?i)(comm|del)")
    
    position_list = []
    
    for m in p.finditer(str_test):
        catching_list.append(m.group().replace("comm","Commencement").replace("del","Delivery"))
        position_list.append(m.span())

    
    return (catching_list,position_list)

def writer_money_regex(str_test):
# For writer, start_quali_condi_regex is different, because there are only "On Commencement" and "On Delivery".

    catching_list = []
    
    p = re.compile(r"(?i)\$\W?([1-9]\d{0,2}[\,\.]?\d{,3}K?M?)")
    
    position_list = []
    
    for m in p.finditer(str_test):
        catching_list.append(m.group())
        position_list.append(m.span())

    
    return (catching_list,position_list)


def writer_comp_spliter(df, comp_id):
    
    
    rows = df[df["Right_COMPENSATION_ID"] == comp_id]
    indexs = rows.index
    
    for ind in indexs:
        print(comp_id)
        row = df.loc[ind]
        desc = row["split_desc"]
        
        
        # Catch Writing_step, money and condition
        
        writing_step_list, step_position_list = writing_step_regex(desc)
        quali_condi_list, quali_position_list = writer_start_quali_condi_regex(desc)
        #payment_type_list, payment_position_list = writer_payment_type_regex(desc)
        writer_money_list, money_position_list = writer_money_regex(desc)
        
        
        # only 1 writing step
        if len(writing_step_list) == 1:
            print("a")
            df.at[ind,"Compensation.Writing_Step"] = writing_step_list[0]
            df.at[ind,"Compensation.Writing_Step_Amount"] = df.loc[ind, 'Right_COMPENSATION_AMOUNT']
            continue
            
        #only 1 quali_condi
        if len(quali_condi_list) == 1 and len(writer_money_list) == 0:
            print("a2")
            
            df.at[ind,"Compensation.Start_Qualifier"] = "On"
            df.at[ind,"Compensation.Payment_Start_Condition"] = quali_condi_list[0]
            continue
            
        # more than 1 writing step and no money
        if len(writing_step_list) > 1 and len(writer_money_list) == 0:
            print("b")
            amount = df.loc[ind, 'Right_COMPENSATION_AMOUNT']
            df = line_spliter(df, ind, len(writing_step_list))
            df.at[ind,"Compensation.Writing_Step"] = writing_step_list
            df.at[ind,"Compensation.Writing_Step_Amount"] = np.concatenate(([amount],
                                                                            np.repeat(0,len(writing_step_list) - 1)))
            continue
            


        # more than 1 pairs of writing step and money, no quali_condi
        if len(writing_step_list) > 1 and len(writer_money_list) > 1 and len(writing_step_list) == len(writer_money_list) \
        and len(quali_condi_list) == 0:
            print("c")
            df = line_spliter(df, ind, len(writing_step_list))
            df.at[ind,"Compensation.Writing_Step"] = writing_step_list
            
            transfered_money_list = []
            for money in writer_money_list:
                transfered_money_list.append(money_transfer(money))
            df.at[ind,"Compensation.Writing_Step_Amount"] = transfered_money_list
            continue
        
        # more than 1 pairs of quali_condi and money, no writing step.
        
        # Do I need to check if sum(writer_money_list) == 'Right_COMPENSATION_AMOUNT'?
        
                    #and sum(writer_money_list) == df.loc[ind, 'Right_COMPENSATION_AMOUNT']
        if len(quali_condi_list) > 1 and len(writer_money_list) > 1 and len(quali_condi_list) == len(writer_money_list) \
        and len(writing_step_list) == 0 :
            print("d")
            df = line_spliter(df, ind, len(quali_condi_list))
            transfered_money_list = []
            for money in writer_money_list:
                transfered_money_list.append(money_transfer(money))
            df.at[ind,"Compensation.%_Amount"] = transfered_money_list 
            df.at[ind,"Compensation.Start_Qualifier"] = np.repeat("On",len(quali_condi_list))
            df.at[ind,"Compensation.Payment_Start_Condition"] = quali_condi_list
            continue
            
        
        
        # there are all writing step, money and quali_condi, and they 3 may be not mathced.
        if len(quali_condi_list) > 1 and len(writer_money_list) > 1 and len(writing_step_list) >1 \
        and len(quali_condi_list) == len(writer_money_list) :
            print('e')
            df = line_spliter(df, ind, len(writer_money_list))
            transfered_money_list = []
            for money in writer_money_list:
                transfered_money_list.append(money_transfer(money))
                
            df.at[ind,"Compensation.%_Amount"] = transfered_money_list
            df.at[ind,"Compensation.Start_Qualifier"] = np.repeat("On",len(quali_condi_list))
            df.at[ind,"Compensation.Payment_Start_Condition"] = quali_condi_list    
            
            
            step_positions = [0]
            
            for step_position in step_position_list:
                step_positions.append(step_position[0])
                step_positions.append(step_position[1])
                
            step_position_list.append(len(desc))
            
            step_lists_repeated = []

            j = 0
            for i in range(0,len(writer_money_list)):
                if money_position_list[i][0] > step_position_list[2*j] and money_position_list[i][1]<step_position_list[2*j+1]:
                    step_lists_repeated.append(writing_step_list[j])
                elif i == 0:
                    continue
                else:
                    j = j + 2
           
            df.at[ind,"Compensation.Writing_Step"] = step_lists_repeated

    
    return df.reset_index(drop=True)


#Main Iteration Function
    
def split_pasting_func(func_df,func):
    df = func_df.copy()
    id_list = df.loc[df["Index"] == "New"]["Right_COMPENSATION_ID"].unique()
    for comp_id in id_list:
        print(func)
        print(comp_id)
        
        if func == "Financier" or func == "Generic_Functions" or func == "Researcher":
            continue
 
        if func == "Writer":
            
            df = overage_spliter(df, comp_id)

            df = Writer_Payment_Type_spliter(df, comp_id)
            
            # Core Part!!! About how to split Compensation_Amount---Writing_Step_Amount---%_Amount
            
            df = writer_comp_spliter(df, comp_id)
            # Core Part!!!
            
            
            df = dura_num_freq_paster(df, comp_id)
            df = rate_num_freq_paster(df, comp_id)
            df = start_date_paster(df, comp_id)
            df = fee_type_paster(df, func, comp_id)

            continue


        if func == "Rights":
            
            df = dura_num_freq_paster(df, comp_id)        
            df = Payment_Type_spliter(df, comp_id)         
            df = start_quali_condi_paster(df, comp_id)           
            df = total_guarant_paster(df, comp_id)
            df = check_number_paster(df, comp_id)
            df = start_expiry_date_paster(df, comp_id)
            df = date_paid_paster(df, comp_id)

            continue
        
        df = comp_type_paster(df, comp_id)
        #df = orig_deal_paster(df, comp_id)
        df = overage_spliter(df, comp_id)
        df = Payment_Type_spliter(df, comp_id)
        
        
        
        # How about combining these two functions together?
        df = per_amount_spliter(df, comp_id)
        df = start_quali_condi_paster(df, comp_id)
        
        
     
        df = dura_num_freq_paster(df, comp_id)
        df = rate_num_freq_paster(df, comp_id)
        df = start_date_paster(df, comp_id)
        
        
        df = fee_type_paster(df, func, comp_id)
        
        if func == "Producer":
            df = service_type_paster(df, func, comp_id)
  
    return df
    
# Main Loop
#Using this dictionary to reorder the output columns.
column_order_list = {'Actor':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Applicable_ind",
                              "Compensation.Compensation_Type","Compensation.Fee_Type",
                              "Compensation.Start_Date","Compensation.Duration_#","Compensation.Duration_Freq",
                              "Compensation.Rate",
                              "Compensation.Rate_Freq","Compensation.Payment_Type","Compensation.%_Total",
                              "Compensation.%_Amount","Compensation.Start_Qualifier","Compensation.Payment_Start_Condition",
                              "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                              "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                              "Right_PROJECT_ID"],
                     'Casting':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Applicable_ind",
                              "Compensation.Compensation_Type","Compensation.Fee_Type",
                              "Compensation.Start_Date",
                              "Compensation.Duration_#","Compensation.Duration_Freq","Compensation.Rate",
                              "Compensation.Rate_Freq","Compensation.Payment_Type","Compensation.%_Total",
                              "Compensation.%_Amount","Compensation.Start_Qualifier","Compensation.Payment_Start_Condition",
                              "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                              "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                              "Right_PROJECT_ID"],
                     'Consultant':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Applicable_ind",
                              "Compensation.Compensation_Type","Compensation.Fee_Type",
                              "Compensation.Start_Date",
                              "Compensation.Duration_#","Compensation.Duration_Freq","Compensation.Rate",
                              "Compensation.Rate_Freq","Compensation.Payment_Type","Compensation.%_Total",
                              "Compensation.%_Amount","Compensation.Start_Qualifier","Compensation.Payment_Start_Condition",
                              "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                              "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                              "Right_PROJECT_ID"],
                     'Director':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                                 "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Applicable_ind",                                 
                              "Compensation.Compensation_Type","Compensation.Fee_Type","Compensation.Start_Date",
                              "Compensation.Duration_#","Compensation.Duration_Freq","Compensation.Rate",
                              "Compensation.Rate_Freq","Compensation.Payment_Type","Compensation.%_Total",
                              "Compensation.%_Amount","Compensation.Start_Qualifier","Compensation.Service_Start_Condition",
                              "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                              "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                              "Right_PROJECT_ID"],
                     'Financier':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE",
                              "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                              "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                              "Right_PROJECT_ID"],
                     'Rights':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Applicable_ind","Compensation.Duration_#",
                               "Compensation.Duration_Freq","Compensation.Payment_Type","Compensation.%_Total",
                               "Compensation.Start_Qualifier","Compensation.Payment_Start_Condition",
                               "Compensatipon.Total_Guaranteed_Commitment","Rights Start Condition",
                               "Rights Start Date","Compensation.Check_#","Compensation Commitment Date","# of Payments",
                               "Frequency","Start_Date","Expiry_Date","Purchase_Date","Reversion_Date",
                               "Date Paid","Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                              "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                              "Right_PROJECT_ID"],
                    'Generic_Functions':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID",
                                         "COMPENSATION_AMOUNT","COMPENSATION_DESC","COMPENSATION_TYPE",
                                          "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                                          "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                                          "Right_PROJECT_ID"],
                     'Producer':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Applicable_ind",
                              "Compensation.Compensation_Type","Compensation.Service_Type","Compensation.Fee_Type",
                              "Compensation.Start_Date",
                              "Compensation.Duration_#","Compensation.Duration_Freq","Compensation.Rate",
                              "Compensation.Rate_Freq","Compensation.Payment_Type","Compensation.%_Total",
                              "Compensation.%_Amount","Compensation.Start_Qualifier","Compensation.Payment_Start_Condition",
                              "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                              "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                              "Right_PROJECT_ID"],
                     'Researcher':["DARTS_DIVISION","FUNCTION","COMPENSATION_ID","DEAL_ID","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Applicable_ind",
                              "Compensation.Compensation_Type","Compensation.Service_Type","Compensation.Fee_Type",
                              "Compensation.Start_Date",
                              "Compensation.Duration_#","Compensation.Duration_Freq","Compensation.Rate",
                              "Compensation.Rate_Freq","Compensation.Payment_Type","Compensation.%_Total",
                              "Compensation.%_Amount",
                               "Compensation.Start_Qualifier",
                               "Index","Compensatipon.Total_Guaranteed_Commitment",
                              "Compensation.Service_Start_Condition","Right_DARTS_DIVISION",
                               "Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                               "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE",
                               "Right_DEAL_ID","Right_FUNCTION","Right_PROJECT_ID"],
                     'Writer':["DARTS_DIVISION","COMPENSATION_ID","DEAL_ID","FUNCTION","PROJECT_ID","COMPENSATION_AMOUNT",
                              "COMPENSATION_DESC","COMPENSATION_TYPE","Compensation.Applicable_ind",
                              "Compensation.Fee_Type",
                              "Compensation.Writing_Step","Compensation.Writing_Step_Amount","Compensation.Start_Date",
                              "Compensation.Duration_#","Compensation.Duration_Freq","Compensation.Rate",
                              "Compensation.Rate_Freq","Compensation.Payment_Type","Compensation.%_Total",
                              "Compensation.%_Amount","Compensation.Start_Qualifier","Compensation.Payment_Start_Condition",
                              "Index","Right_DARTS_DIVISION","Right_COMPENSATION_ID","Right_COMPENSATION_AMOUNT",
                              "Right_COMPENSATION_DESC","Right_COMPENSATION_TYPE","Right_DEAL_ID","Right_FUNCTION",
                              "Right_PROJECT_ID"]}
                     
                     
# Main loop
for func in dict_combined.keys():
    dict_combined[func] = split_pasting_func(dict_combined[func],func)
    #drop the temporary column
    dict_combined[func] = dict_combined[func].drop(columns = ["split_desc"])
    #One more step, choose excat columns with exact order.
    dict_combined[func] = dict_combined[func][column_order_list[func]]
    
    
#Writing excel
for func in dict_combined.keys():
    with pd.ExcelWriter(output_position + '/' + date + "_Fixed_" + func + "_" +
                              "Parsed.xlsx", engine='xlsxwriter') as writer:
        dict_combined[func].to_excel(writer, sheet_name='Fixed_Compensation',index = False)
        
print("Parsing complete.")































    
