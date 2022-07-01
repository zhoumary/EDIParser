"""
1. read original MIG file
2. create a excel file
3. set the first sheet header
4. invoke parser module to parse the MIG
"""

from turtle import clear
import pandas
import os
import xlsxwriter

# read the data from the path /data/MIG52eOriginal.xlsx
absolutepath = os.path.abspath(__file__)
print(absolutepath)

fileDirectory = os.path.dirname(absolutepath)
print(fileDirectory)
# Path of parent directory
parentDirectory = os.path.dirname(fileDirectory)
print(parentDirectory)
# Navigate to data directory
newPath = os.path.join(parentDirectory, 'data')   
print(newPath)

try:
    file = "\data\MIG52eOriginal.xlsx"
    path = os.getcwd()+file
    print(path)

    """
    1. parsing the original MIG to Dictionary with the first row as key
    2. according to the EDI_MIG_HEADER.json to parse the row data
        2.1 Level0/1/2/3/4 <= Ebene + Bez
        2.2 Content(temparary null, which need to parsing from the MIG detailed content)
        2.3 Repeat Times <= MaxWdhBDEW
        2.4 Content Type <= Bez with leading "SG"
        2.5 Desc. <= Inhalt
    """    
    mig_original = pandas.read_excel(path)
    mig_original_dict = mig_original.to_dict('records')
    
    mig_hierarchy = []
    mig_hierarchy_dict = {'Level 0': "", 'Level 1': "", 'Level 2': "", 'Level 3': "", 'Level 4': "", 'Content': "", 'Repeat Times': "", 'Content Type': "", 'Desc.': ""}

    for index in range(len(mig_original_dict)):

        # 2.1 Level0/1/2/3/4 <= Ebene + Bez
        mig_layer = mig_original_dict[index]["Ebene"]
        mig_layer_content = mig_original_dict[index]["Bez"]
        match mig_layer:
            case 0: 
                mig_hierarchy_dict["Level 0"] = mig_layer_content
            case 1:
                mig_hierarchy_dict["Level 1"] = mig_layer_content
            case 2:
                mig_hierarchy_dict["Level 2"] = mig_layer_content
            case 3:
                mig_hierarchy_dict["Level 3"] = mig_layer_content
            case 4:
                mig_hierarchy_dict["Level 4"] = mig_layer_content
            case _:
                print("Code not found")
        
        # 2.2 Content <= from the AHB document, temparory null


        # 2.3 Repeat Times <= MaxWdhBDEW
        mig_repeat_times = mig_original_dict[index]["MaxWdhBDEW"]
        mig_hierarchy_dict["Repeat Times"] = mig_repeat_times

        # 2.4 Content Type: Group or Element <= Bez with leading "SG"
        if mig_layer_content.startswith("SG"):
            mig_type = "Group"
        else:
            mig_type = "Element"
        mig_hierarchy_dict["Content Type"] = mig_type

        # 2.5 Desc. <= Inhalt
        mig_desc = mig_original_dict[index]["Inhalt"]
        mig_hierarchy_dict["Desc."] = mig_desc

        mig_hierarchy.append(mig_hierarchy_dict)

        print(mig_hierarchy_dict)
        print(mig_hierarchy_dict.items())
        for mig_key, mig_value in mig_hierarchy_dict.items():
            if isinstance(mig_hierarchy_dict[mig_key], str):
                mig_hierarchy_dict[mig_key] = ""
            else:
                mig_hierarchy_dict[mig_key] = 0
    
except FileNotFoundError:
    print("Please check the path.")

# create an excel file to save the data parsed
workbook = xlsxwriter.Workbook('UTILMD MIG 5.2e.xlsx')
worksheet = workbook.add_worksheet()
# write the sheet header
sheetheader = ['Level 0', 'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Content', 'Repeat Times', 'Content Type', 'Desc.']
# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0
# Iterate over the data and write it out row by row.
for item in (sheetheader):
    worksheet.write(row, col + 1, item)

# write the row data with mig_hierarchy


workbook.close()
