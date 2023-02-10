import os
import linecache
import xlwt
from xlwt import Workbook
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

n = 27
while n < 28:
    s = str(n)
    os.system('cmd /c "color a & systeminfo /s 153.90.162.' + s + '|findstr /i \"host OS\" > e:\output' + s + '.txt"')
    #os.system('cmd /c "color a & systeminfo /s 153.90.162.' + s + '|findstr /i \"host OS\" > e:\output.txt"')
    nameos = linecache.getline(r"E:\output" + s + ".txt", 2)
    nameos = nameos.lstrip("OS Name:                   ")
    versionOS = linecache.getline(r"E:\output" + s + ".txt", 3)
    versionOS = versionOS.lstrip("OS Version:                   ")
    sheet1.write(n-1,1, nameos)
    sheet1.write(n-1,2, versionOS)
    sheet1.write(n-1,0, "153.90.162." + s)
    os.remove("E:\output" + s + ".txt")
    n += 1


wb.save('computerOS.xls')
