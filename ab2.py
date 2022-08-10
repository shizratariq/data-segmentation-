from openpyxl import load_workbook
import pyexcel as pe
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
for i in range (8):
    k, l = segment(k, l)
    print(k, " ", l)
    print (list2d[i])
pe.save_as(array=list2d, dest_file_name="example.xlsx")
# for i in range(k, len(res)):
#     if res[i] > 100:
#         seg1.append(res[i])
#         list2d[l].append(res[i])
#     elif len(seg1) > 40:
#         end=i+2
#         k=end
#         l=l+1
#         seg1 = []
#         break
# print (list2d[0])
# print (k)
# print (l)
# seg1=[]
# for i in range(k, len(res)):
#     if res[i] > 100:
#         seg1.append(res[i])
#         l = 1
#         list2d[l].append(res[i])
#     elif len(seg1) > 40:
#         end=i
#         l=l+1
#         break
# print (list2d[1])
#
# end= segment(k,l)
# write(0)
# print(list2d[0])
# l=l+1
# end= segment(end,l)
# write(0)
# print(list2d[1])