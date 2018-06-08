### Reading exported CSV and selecting only required columns into a XLS file
# Import modules 
import pandas as pd 
import os
import sys

# Use sys functions to change encoding to enable csv to xls format
reload(sys)
sys.setdefaultencoding('utf8')

# Establish original csv to read from: "base"
base = pd.read_csv('Autodesk Arnold Technical Resources 6_8_2018-06-08_124704_output.csv')

# Establish range of columns needed
length = len(base.columns)
pos = [0, 1, 2, 9, 10, 19, 28]
pos2 = range(35, length)
colnames = pos + pos2

final = pd.read_csv('Autodesk Arnold Technical Resources 6_8_2018-06-08_124704_output.csv', usecols = colnames)

# Filter and then create new dataframe to take filter as argument
reachable_filter = (final["MEI"] >= 25) & (final["active"] == True) 
reachable_df = final[reachable_filter]

not_reachable_filter = (final["MEI"] < 25) & (final["active"] == True)
not_reachable_df = final[not_reachable_filter]

# Specify which parameters need to be satisfied
no_match_filter = (final["active"] == False)
no_match_df = final[no_match_filter] 

# Set new file path/directory for newly created xls files
path = "/Users/jshek/Desktop/completed xls"
os.chdir(path)

from pandas import ExcelWriter
writer = pd.ExcelWriter('Autodesk Arnold Technical Resources 6_8_2018-06-08_124704_output.xlsx', engine = "xlsxwriter")

reachable_df.to_excel(writer, sheet_name = 'Reachable', index = False)
not_reachable_df.to_excel(writer, sheet_name = 'Not Reachable', index = False)
no_match_df.to_excel(writer, sheet_name = 'No Match', index = False)



