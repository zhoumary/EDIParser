"""
1. read original MIG file
2. create a excel file
3. set the first sheet header
4. invoke parser module to parse the MIG
"""

import pandas
import os
import xlsxwriter

# upload the word file of 5.2e MIG.docx
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
        2.2 Content(temparary null, which need to parsing from the word/pdf)
        2.3 Repeat Times <= MaxWdhBDEW
        2.4 Content Type <= Bez with leading "SG"
        2.5 Desc. <= Inhalt
    """    
    mig_original = pandas.read_excel(path)
    mig_original_dict = mig_original.to_dict('records')
    
    mig_keys = ['Level 0', 'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Content', 'Repeat Times', 'Content Type', 'Desc.']

    for index in range(len(mig_original_dict)):
        print(mig_original_dict[index])
        # 2.1 Level0/1/2/3/4 <= Ebene + Bez
        mig_layer = mig_original_dict[index]["Ebene"]
        print(mig_layer)
        # 2.3 Repeat Times <= MaxWdhBDEW
        mig_repeat_times = mig_original_dict[index]["MaxWdhBDEW"]
        print(mig_repeat_times)
        # 2.4 Content Type <= Bez with leading "SG"

        # 2.5 Desc. <= Inhalt
        mig_desc = mig_original_dict[index]["Inhalt"]
        print(mig_desc)

    
except FileNotFoundError:
    print("Please check the path.")

# # create an excel file to save the data parsed
# workbook = xlsxwriter.Workbook('UTILMD MIG 5.2e.xlsx')
# worksheet = workbook.add_worksheet()

# workbook.close()
