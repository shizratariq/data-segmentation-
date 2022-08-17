from openpyxl import load_workbook
import pyexcel as pe
import numpy as np
book=load_workbook('power.xlsx')
sheet=book['Trial1']

#initializing variables
k=0
res=[]
fsr1=[]
seg1=[]
end=0
l=0
list2d=[[] for i in range(30)]
#removing string value from array
for row in sheet.rows:
    str(row[0].value)
    fsr1.append(row[0].value)
fsr1.pop(0)
#removing none value from array
for val in fsr1:
    if val != None :
        res.append(val)
l = 0
def segment(k,l):
    seg1 = []
    for i in range(k, len(res)):
        if res[i] > 100:
            seg1.append(res[i])
            list2d[l].append(res[i])
        elif len(seg1) > 40:
            end = i + 1
            k = end
            l = l + 1
            return k, l
            break
final30=[[] for i in range(30)]
m=0
for i in range (8):
    k, l = segment(k, l)
    print(k, " ", l)
    mid = 0 + (len(list2d[i]) - 0) / 2
    print (mid)
    mid=int(mid)
    start30= mid-15
    end30=mid+15
    for j in range (len(list2d[i])):
        if j>start30-1 and j<end30:
            final30[m].append(list2d[i][j])
    m=m+1
    print (final30[i])
    print (list2d[i])


pe.save_as(array=final30, dest_file_name="example.xlsx")
