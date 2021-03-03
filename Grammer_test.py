import openpyxl
import pandas as pd
import os

def list_files():
    file_data = []
    i=1
    for file in os.listdir("./"):
        if file.endswith(".xlsx"):
            print(str(i)+'. ', file)
            file_data.append(file)
            i=i+1
    return  file_data

file_data = list_files()
file_num = int(input("please enter no. of files: "))

file = file_data[file_num-1]
print(file)


"""
wb = openpyxl.load_workbook('data.xlsx')
sheet = wb.get_sheet_by_name('Sheet2')
range = ['A3':'D20']   #<-- how to specify this?
spots = pd.DataFrame(sheet.range) #what should be the exact syntax for this?

print (spots)
"""