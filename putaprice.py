import numpy as np
import pandas as pd

import os

#search code
def search_code(code, data):
    check = data[data.isin([code]).any(axis= 1)]
    return check

#find header
def find_header(data):
    header = pd.DataFrame()
    column_names = data.columns
    for column_name in column_names:
        header = pd.concat([header, data[data[column_name].str.contains("НДС")]])
    header = header.iloc[0].to_list()
    header = [str(elem) for elem in header]
    header = unique_header(header)
    #print(header)
    return header

#make unique header
def unique_header(header):
    un_header = []
    for name in header:
        if name not in un_header:
            un_header.append(name)
        else:
            un_header.append(name + "1")
    return un_header

#check FEE
def check_fee(header):
    fee_flag = 0.8
    for elem in header:
        if "НДС" in elem:
            if "без НДС" in elem:
                fee_flag = 1
            break
    return fee_flag, elem

#function to search price
def search_price(data, fee_column, fee_value):
    price = 0
    price = float(data[fee_column].iloc[0])*float(fee_value) 
    return price

#search for excel
path = os.getcwd()
list_of_files = os.listdir(path)
list_of_data = []
for elem in list_of_files:
    if elem.split(".")[-1] >= "xl":
        if elem.split(".")[0] != "target":
            data = pd.read_excel(elem, dtype = "string")
            header = find_header(data)
            data.columns = header
            list_of_data.append(data)

#collect data of target and quantity
target_data = pd.read_excel("target.xlsx", header = None, dtype="string")
target_data.columns = ["code","quantity"]
# print(target_data)

#print(list_of_data[0]["Тариф \n14.08.2023, Руб, \nза единицу тарифа, \nбез НДС"])
#print(list_of_data[1]["Тариф с НДС, руб"])
#print(list_of_data[2]["Цена с НДС, руб./м(шт)"])

#loop
price_list = [0.0 for elem in target_data["code"]]
for data in list_of_data:
    fee_value, fee_column = check_fee(data)
    for pos, code in enumerate(target_data["code"]):
        code_exists = search_code(code, data)
        if code_exists.empty != True:
            if price_list[pos] == 0.0:
                price = search_price(code_exists, fee_column, fee_value)
                price_list[pos] = (float(price))
                print(str(pos) + str(price))

#write results
target_data.insert(2, "Price", price_list, True)
cost_sum = []
for quantity, price in zip(target_data["quantity"], price_list):
    cost_sum.append(str(float(quantity) * float(price)))
target_data.insert(3, "Cost", cost_sum, True)
print(target_data)
target_data.to_excel("output.xlsx")