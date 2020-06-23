import os
import sys
import xlsxwriter
import re
row =0
col=0
c=0
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Workbook.xlsx',{'strings_to_numbers': True})
path = ""   # path to your files you want to read
regex = "[a-z]"
all_files = os.listdir(path) 
for file in all_files:
    c=c+1
    with open(os.path.join(path, file)) as f:
      if(file.endswith('.txt')):    #you can add any file extension that needs to be written
#This is to reduce the length of file name
        l = file.split("_")
        if(len(l[4]) > 27):
          l[4] =l[4][0:27]
          l[4] =l[4]+str(c)
        else:
          l[4] = l[4]+str(c)
        name = workbook.add_worksheet(l[4]) 
        row = 0         
        for line in f:
          line = line[:-1]
          match = re.search((regex),line)
          if(match == None):
            value =list(line.split(" "))
            value = [x for x in value if x !=""]
            col =0
            row = row+1
            for i in range(len(value)):
                name.write(row, col, (value[i]))
                col = col+1
workbook.close() 