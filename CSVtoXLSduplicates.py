### Reading exported CSV and selecting only required columns into a XLS file
# Import modules 
import pandas as pd
import os 
import sys

# Use sys functions to change encoding to enable csv to xls format
reload(sys)
sys.setdefaultencoding('utf8')

# Establish original csv to read from: "base"
base = pd.read_csv('Cisco NVS File 6_11_2018-06-11_124799_output.csv')

# Establish range of columns needed - will need to keep Match Type and Active status for filters 
length = len(base.columns)
pos = [0, 1, 2, 8, 9, 10, 19, 28]
# From column 35 on is all custom attributes, so must include all columns that appear
pos2 = range(35, length)
colnames = pos + pos2

final = pd.read_csv('Cisco NVS File 6_11_2018-06-11_124799_output.csv', usecols = colnames)

# Filter and then create new dataframe to create new tab in xls 
reachable_filter = (final["MEI"] >= 25) & (final["active"] == True) 
reachable_df = final[reachable_filter]
reachable_sort = reachable_df.sort_values("MEI", inplace = False)

not_reachable_filter = (final["MEI"] < 25) & (final["active"] == True)
not_reachable_df = final[not_reachable_filter]
not_reachable_sort = not_reachable_df.sort_values("MEI", inplace = False, ascending = False)

# No matches can also be duplicates so need to figure out a way that eliminates the duplicate rows from no match tab
no_match_filter = (final["active"] == False) 
temp_df = final[no_match_filter]
# Filter out two criteria from dataframe by placing both criteria in () brackets 
no_match_df = temp_df[(temp_df["Match Type"] != "duplicate match") & (temp_df["Match Type"] != "duplicate input")]
# no_match_sort = no_match_df.sort_values("MEI", inplace = False, ascending = False)

# Created this to place duplicate results in duplicates tab
duplicate_filter = (final["active"] == False) & (final["Match Type"] == "duplicate match") | (final["Match Type"] == "duplicate input")
duplicate_df = final[duplicate_filter]
# duplicate_sort = duplicate_df.sort_values("MEI", inplace = False, ascending = False) 

# Set new file path/directory for newly created xls files
os.chdir("/Users/jshek/Desktop/completed xls")

from pandas import ExcelWriter
writer = pd.ExcelWriter('Cisco NVS File 6_11_2018-06-11_124799_output.xlsx', engine = "xlsxwriter")

reachable_sort.to_excel(writer, sheet_name = 'Reachable', index = False)
not_reachable_sort.to_excel(writer, sheet_name = 'Not Reachable', index = False)
no_match_df.to_excel(writer, sheet_name = 'No Match', index = False)
duplicate_df.to_excel(writer, sheet_name = 'Duplicates', index = False)
