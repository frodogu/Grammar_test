# - List all files in current directory and select one to read from
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

# - Read from selected file and choose selection of range
import pandas as pd
sheet = pd.read_excel(file, sheet_name='New Hire')
sheet = sheet.iloc[2:, :34]

sheet.dropna(subset=["SAP ID(员工的唯一标识)"],inplace=True)
print(sheet)
