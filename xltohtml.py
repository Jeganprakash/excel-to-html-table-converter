import pandas as pd
import sys
from xlrd import XLRDError

input=[]
try:
    excel=sys.argv[1]
except IndexError:
    print('''provide .xlsx filename as commandline argument
run this command:
    $python convertable.py example.xlsx ''')
    exit(0)
#reading excel sheet from commandline
try:
    a=pd.read_excel(excel,sheet_name=0)
except XLRDError as e:
    print(e)
    print("please only provide .xlsx file")
    exit(0)
except FileNotFoundError:
    print("File not found.....\n please check .xlsx file is in the current directory or provide correct directory/filename")
    exit(0)

rows=a.shape[0]             # no of rows
columns=a.shape[1]          # no of columns           
header_values=list(a.head(0)) # table headings column

#list to string
def listToString(s):  
    str1 = ""     
    for ele in s:  
        str1 += ele   
    return str1  


table='''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>HTML Table</title>
</head>
<body>\n<table>\n<thead class="tablehead">\n<tr class="headerrow">\n'''
#table head
for i in header_values:
    input.append('\t<th class="thd"><strong>{0}</strong></th>\n'.format(i))

table=table + listToString(input)+ '</tr>\n</thead>\n<tbody class="tablebody">\n'

#table rows
input=[]
for i in range(rows):
    row_value=a.values[i]
    input.append('<tr class="tabrows trow{}">\n'.format(i))
    for j in range(columns):
        input.append('\t<td class="tdata td{1}">{0}</td>\n'.format(row_value[j],j))
    input.append('</tr>\n')
 

table=table+listToString(input)+'</tbody>\n</table>\n</body>\n</html>'


#output file
output_file=open("{}.html".format(excel),"w")
output_file.write(table)
output_file.close()

print("success! \n {}.html file with table is created in the current directory".format(excel))