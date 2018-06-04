#!/usr/bin/python

import functions
import xlrd
import xlwt
import os





path = r"/var/www/html/uploads/dimtable.xls"
os.chmod(path,0o777)
wb = xlwt.Workbook()
outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)
insheet = xlrd.open_workbook(path).sheets()[0]
for row in range(0, insheet.nrows):
    for col in range(0, insheet.ncols):
        outsheet.write(row, col, insheet.cell_value(row, col))
lis=[]
for cols in range(0,insheet.ncols):
    lis.append(insheet.cell_value(0,cols))
if "Short description" in lis:
    ind=lis.index("Short description")
if "description" in lis:
    ind = lis.index("description")
outsheet.write(0,ind,"description")
for rows in range(1,insheet.nrows):
    outsheet.write(rows,ind,None)
wb.save(path)
functions.dim_table(path)


print("Content-type:text/html\r\n\r\n")
print('<html><body><h2><center>uploaded to Dataware as "Dim_incident"</center></h2></body></html>')
