from openpyxl import load_workbook
# df = pd.read_excel('power.xlsx', sheet_name = [1,2,3,4,5])
# # df.keys()
# # print(df)
import pyexcel as pe
import pandas as pd
import os
import glob
filename=[]
# use glob to get all the csv files
# in the folder
path = os.getcwd()
csv_files = glob.glob(os.path.join(path, "*.xlsx"))
for f in csv_files:
    # read the csv file
    df = pd.read_excel(f)
    name=f.split("\\")[-1]
    filename.append(name)
    # print('File Name:', f.split("\\")[-1])
print(type(name))
for i in range (len(filename)):
    print(filename[i])

tabs = pd.ExcelFile(filename[2]).sheet_names
print(tabs)
book=load_workbook(filename[2])
sheet=book[tabs[1]]

def repfunc():
    # initializing variables
    k = 0
    res = []
    fsr1 = []
    fsr2 = []
    fsr3 = []
    fsr4 = []
    seg1 = []
    end = 0
    l = 0
    s = 0
    list2d = [[] for i in range(30)]
    final30 = [[] for i in range(30)]
    m = 0
    finalfsr2 = [[] for i in range(30)]
    finalfsr3 = [[] for i in range(30)]
    finalfsr4 = [[] for i in range(30)]
    def rem_str_value():
        # removing string value from array for FSR1
        for row in sheet.rows:
            str(row[0].value)
            fsr1.append(row[0].value)
        fsr1.pop(0)
        # removing string value from array for FSR2
        for row in sheet.rows:
            str(row[1].value)
            fsr2.append(row[1].value)
        fsr2.pop(0)
        # removing string value from array for FSR3
        for row in sheet.rows:
            str(row[2].value)
            fsr3.append(row[2].value)
        fsr3.pop(0)
        # removing string value from array for FSR4
        for row in sheet.rows:
            str(row[3].value)
            fsr4.append(row[3].value)
        fsr4.pop(0)
        # removing none value from array
        for val in fsr1:
            if val != None:
                res.append(val)

    def segment(k, l, s):
        seg1 = []
        # getting start point location of FSR1
        for i in range(k, len(res)):
            if res[i] > 100:
                s = i + 2
                # print (s)
                break
        # getting start point location of FSR1
        for i in range(k, len(res)):
            if res[i] > 100:
                seg1.append(res[i])
                list2d[l].append(res[i])
            elif len(seg1) > 40:
                end = i + 1
                k = end
                l = l + 1
                return k, l, s
                break

    def fsrsegment(s, l):
        for i in range(s, s + 30):
            finalfsr2[l].append(fsr2[s - 2])
            finalfsr3[l].append(fsr3[s - 2])
            finalfsr4[l].append(fsr4[s - 2])
            s = s + 1

    rem_str_value()
    for i in range(8):
        # k=ending point
        # l= number of segment
        # s is the starting point
        k, l, s = segment(k, l, s)
        print(k, " ", l)
        mid = 0 + (len(list2d[i]) - 0) / 2
        # print (mid)
        mid = int(mid)
        start30 = mid - 15
        end30 = mid + 15
        s = s + start30
        print(s)
        for j in range(len(list2d[i])):
            if j > start30 - 1 and j < end30:
                final30[m].append(list2d[i][j])
        m = m + 1
        fsrsegment(s, l)
        print(final30[i])
        print(finalfsr2[l])
        print(finalfsr3[l])
        print(finalfsr4[l])
repfunc()