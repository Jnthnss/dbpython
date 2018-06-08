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

# Establish range of columns needed - in this script will need to keep match type, column #8
length = len(base.columns)
pos = [0, 1, 2, 8, 9, 10, 19, 28]
pos2 = range(35, length)
colnames = pos + pos2

final = pd.read_csv('Autodesk Arnold Technical Resources 6_8_2018-06-08_124704_output.csv', usecols = colnames)

# Filter and then create new dataframe to create new tab in xls 
reachable_filter = (final["MEI"] >= 25) & (final["active"] == True) 
reachable_df = final[reachable_filter]

not_reachable_filter = (final["MEI"] < 25) & (final["active"] == True)
not_reachable_df = final[not_reachable_filter]

# Specify which parameters need to be satisfied
no_match_filter = (final["active"] == False) 
no_match_df = final[no_match_filter] 

# Created this to remove duplicate SIDs 
duplicate_filter = (final["active"] == False) & (final["Match Type"] == "duplicate match") | (final["Match Type"] == "duplicate input")
duplicate_df = final[duplicate_filter] 

# Set new file path/directory for newly created xls files
path = "/Users/jshek/Desktop/completed xls"
os.chdir(path)

from pandas import ExcelWriter
writer = pd.ExcelWriter('Autodesk Arnold Technical Resources 6_8_2018-06-08_124704_output.xlsx', engine = "xlsxwriter")

reachable_df.to_excel(writer, sheet_name = 'Reachable', columns = excols, index = False)
not_reachable_df.to_excel(writer, sheet_name = 'Not Reachable', index = False)
no_match_df.to_excel(writer, sheet_name = 'No Match', index = False)
duplicate_df.to_excel(writer, sheet_name = 'Duplicates', index = False)

# Remove unnecessary columns that show up on exported csv 


# Variable to remove extra columns from exported csv - change 8 depending on how many custom attribute columns there are

