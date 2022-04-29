"""
1. create a excel file
2. set the first sheet header
3. invoke parser module to parse the MIG
"""

import xlsxwriter

# upload the word file of UTILMDMIG.doc


# create an excel file to save the data parsed
workbook = xlsxwriter.Workbook('UTILMD MIG 5.2e.xlsx')
worksheet = workbook.add_worksheet()



workbook.close()
