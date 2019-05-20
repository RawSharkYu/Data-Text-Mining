
"""
Created on Tue Feb 19 09:46:13 2019

@author: AYu5
"""
import sys
import pandas as pd
import numpy as np
import string
import re
 

# Save "Delta_Payment_Termdeal_Existed.xlsx" as 
# "Delta_Payment_Termdeal_Existed.csv"

# Change following 3 variables into current positions and date


input_position = (r'C:\Users\ayu5\Desktop\Old_Data_Parsing\0409DataParsing\raw data\04062019 - DATA\Payment\Payment')
output_position = (r'C:\Users\ayu5\Desktop\0418Updated\output')
date =  '04092019'



# Don't change all the codes below unless necessary and certain

input_position = input_position.replace('\\','/')
output_position = output_position.replace('\\','/')

function_name_list = ['Actor','Casting','Consultant','Director','Financier',
                 'Generic_Functions','Producer','Researcher',
                 'Rights','Termdeal','Writer']

# The matching function
def pay_function(existed_df, new_df):
    
    combined = pd.concat([existed_df,new_df],sort=False,ignore_index=True)
    
    for index, row in combined.iterrows():
        if row['Index'] in ['New',"Updated"]:
            print((index,row['Index']))
            if pd.notna(row['Right_PAYMENT_COMMENTS']):
                comment = row['Right_PAYMENT_COMMENTS']
                comment = re.sub(r'(?i)invo?i?c?e?s?\W*(\d*)\W{,4}.+?(\d*)?','',comment)
                number = re.findall(r"(D?\d{3,}\-?\d{2,})",comment)
                if len(number) > 0:
                    print(number)
                    new_check_number = ' & '.join(number)
                    new_check_number = new_check_number.replace('-',"")
                    print(new_check_number)
                    combined.loc[index, 'Payment.check_number'] = new_check_number
    
    return combined



# Create Empty DataFrame to store combined data
comb_df = pd.DataFrame(columns = ["DARTS_DIVISION","COMPENSATION_ID","PAYMENT_ID","PAYMENT_AMOUNT",
                     "PAYMENT_COMMENTS","PAYMENT_DATE","FUNCTION","Payment.check_number",
                     "Index","Right_DARTS_DIVISION","Right_PAYMENT_ID","Right_COMPENSATION_ID",
                     "Right_PAYMENT_AMOUNT","Right_PAYMENT_COMMENTS","Right_PAYMENT_DATE",
                     "Right_FUNCTION"])

# Main Loop
for func in function_name_list:
    print(func)
    existed_file = "Delta_Payment_" + func + "_Existed.csv"
    new_file = "Delta_Payment_" + func + "_New.csv"
    existed_df = pd.read_csv(input_position + '/' + existed_file)
    new_df = pd.read_csv(input_position + '/' + new_file)
    
    
    existed_df = existed_df.drop(columns=['Index1'])
    new_df = new_df.rename(index=str, 
               columns={'DARTS_DIVISION':'Right_DARTS_DIVISION','COMPENSATION_ID':'Right_COMPENSATION_ID',
                        'FUNCTION':'Right_FUNCTION','PAYMENT_ID':'Right_PAYMENT_ID',
                        'PAYMENT_AMOUNT':'Right_PAYMENT_AMOUNT','PAYMENT_COMMENTS':'Right_PAYMENT_COMMENTS',
                        'PAYMENT_DATE':'Right_PAYMENT_DATE'})
    

    
    # matching function below
    pay_df = pay_function(existed_df, new_df)
    # matching function above

    # Change the order of columns
    pay_df["PAYMENT_DATE"] = pd.to_datetime(pay_df["PAYMENT_DATE"])
    pay_df["Right_PAYMENT_DATE"] = pd.to_datetime(pay_df["Right_PAYMENT_DATE"])
    
    pay_df = pay_df[["DARTS_DIVISION","COMPENSATION_ID","PAYMENT_ID","PAYMENT_AMOUNT",
                     "PAYMENT_COMMENTS","PAYMENT_DATE","FUNCTION","Payment.check_number",
                     "Index","Right_DARTS_DIVISION","Right_PAYMENT_ID","Right_COMPENSATION_ID",
                     "Right_PAYMENT_AMOUNT","Right_PAYMENT_COMMENTS","Right_PAYMENT_DATE",
                     "Right_FUNCTION"]]
    
    comb_df = comb_df.append(pay_df).reset_index(drop=True)
    
    with pd.ExcelWriter(output_position + '/' + "Payment_" + func + "_" +
                                  "Parsed.xlsx", engine='xlsxwriter') as writer:
        pay_df.to_excel(writer, sheet_name='Payment',index = False)

    
with pd.ExcelWriter(output_position + '/' + "Payment_Append.xlsx",
                    engine='xlsxwriter') as writer:
    comb_df.to_excel(writer, sheet_name='Payment',index = False)


