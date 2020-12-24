import sys
import tkinter as tk
from tkinter import filedialog, Text
import tkinter.font as font
from tkinter.filedialog import asksaveasfile

import pandas as pd
import datetime
from openpyxl import Workbook

root = tk.Tk()
root.title('Power Meter Consumption Calculator (GSFA)')
myFont = font.Font(size=8)
rootloc = ''
dfLinks = pd.DataFrame(columns=['Column ', 'Row', 'Number', 'State', 'realColumn'])
dfF = pd.DataFrame(columns=['Hour', 'Sum Week', 'Sum Sat', 'Sum Sun', 'Wint Week', 'Wint Sat', 'Wint Sun'])
monthList = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
dfTotals = pd.DataFrame(columns=['Month', 'Off-peak', 'Standard', 'Peak', 'Max kVA'])
btn = []


def func(name):
    print(name)

def test():
    print(winterP.get())

def changeState(btn, nbr, column, row):
    if btn[nbr]['text'] == 'S':
        btn[nbr].config(text="O", fg="White", bg="Green")
    elif btn[nbr]['text'] == 'O':
        btn[nbr].config(text="P", fg="White", bg="Red")
    elif btn[nbr]['text'] == 'P':
        btn[nbr].config(text="S", fg="Black", bg="Yellow")
    return


def selectIn():
    global rootloc
    filename = filedialog.askopenfilename(initialdir="/", title="Select File",
                                          filetypes=(("csv file ", "*.csv"), ("all files", "*.*")))
    rootloc = filename
    return


def exit():
    sys.exit()
    return


def genOutput():
    # for x, item in enumerate(monthList):
    #     dfTotals.loc[x] = [item] + [0] + [0] + [0] + [0]
    #
    # maxVA = 0
    # up = 2
    # float1 = 0
    # startFlag = 0
    # csvFile = open(rootloc)
    # dateSplit = [2019,1,1]
    # lines = csvFile.readlines()
    # for p in lines:
    #     p1 = p.replace('"', '')
    #     p2 = p1.split(',')
    #
    #     if (p.find('Date') != -1) and (p.find('Total VA') != -1):
    #         startFlag = 1
    #
    #     if startFlag == 1:
    #         try:
    #             dateI = p2.index('Date')
    #             totalI = p2.index('Total VA')
    #             startI = p2.index('Start')
    #             endI = p2.index('End')
    #             importI = p2.index('Import W')
    #         except ValueError:
    #             number = -1
    #             for x, item in enumerate(monthList):
    #                 if p.find(monthList[x]) != -1:
    #                     number = p.find(monthList[x])
    #
    #             if number != -1:
    #                 up = up + 1
    #                 if monthList[dateSplit[1] - 1] != p2[dateI].split('-')[1]:
    #                     maxVA = 0
    #                 else:
    #                     pass
    #                 dateSplit = p2[dateI].split('-')
    #                 dateInteger = monthList.index(dateSplit[1]) + 1
    #                 dateSplit[1] = dateInteger
    #                 dateSplit[2] = '20' + dateSplit[2]
    #                 startHour = int(p2[startI].split(":")[0])
    #                 day = datetime.datetime(int(dateSplit[2]), dateSplit[1], int(dateSplit[0])).weekday()
    #                 float1 += float(p2[importI])
    #                 if (float(p2[totalI]) > maxVA):
    #                     dfTotals.loc[(dateSplit[1] - 1), 'Max kVA'] = float(p2[totalI])
    #                     maxVA = float(p2[totalI])
    #                 if (dateInteger < 4) or (dateInteger > 8):
    #                     if (day < 5):
    #                         if (dfF.iloc[startHour]['Sum Week'] == 'O'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Off-peak'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Sum Week'] == 'S'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Standard'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Sum Week'] == 'P'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Peak'] += float(p2[importI]) / 2
    #                     elif (day == 5):
    #                         if (dfF.iloc[startHour]['Sum Sat'] == 'O'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Off-peak'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Sum Sat'] == 'S'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Standard'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Sum Sat'] == 'P'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Peak'] += float(p2[importI]) / 2
    #                     elif (day == 6):
    #                         if (dfF.iloc[startHour]['Sum Sun'] == 'O'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Off-peak'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Sum Sun'] == 'S'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Standard'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Sum Sun'] == 'P'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Peak'] += float(p2[importI]) / 2
    #                 elif (dateInteger > 3) and (dateInteger < 9):
    #                     if (day < 5):
    #                         if (dfF.iloc[startHour]['Wint Week'] == 'O'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Off-peak'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Wint Week'] == 'S'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Standard'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Wint Week'] == 'P'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Peak'] += float(p2[importI]) / 2
    #                     elif (day == 5):
    #                         if (dfF.iloc[startHour]['Wint Sat'] == 'O'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Off-peak'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Wint Sat'] == 'S'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Standard'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Wint Sat'] == 'P'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Peak'] += float(p2[importI]) / 2
    #                     elif (day == 6):
    #                         if (dfF.iloc[startHour]['Wint Sun'] == 'O'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Off-peak'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Wint Sun'] == 'S'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Standard'] += float(p2[importI]) / 2
    #                         elif (dfF.iloc[startHour]['Wint Sun'] == 'P'):
    #                             dfTotals.loc[(dateSplit[1] - 1), 'Peak'] += float(p2[importI]) / 2
    #
    # for x, row in dfTotals.iterrows():
    #     dfTotals.iloc[x]["Off-peak"] = "{: .2f}".format(dfTotals.iloc[x]["Off-peak"])
    #     dfTotals.iloc[x]["Peak"] = "{: .2f}".format(dfTotals.iloc[x]["Peak"])
    #     dfTotals.iloc[x]["Standard"] = "{: .2f}".format(dfTotals.iloc[x]["Standard"])
    #     dfTotals.iloc[x]["'Max kVA'"] = "{: .2f}".format(dfTotals.iloc[x]["Max kVA"])
    #
    # file = asksaveasfile(initialdir="/", title="Select output location",
    #                                       filetypes=(("csv file ", "*.csv"),("all files", "*.*")), defaultextension = ".csv")
    #
    # dfTotalsString = dfTotals.round(1).astype(str)
    # dfTotalsString['Off-peak'] += " kWh"
    # dfTotalsString['Standard'] += " kWh"
    # dfTotalsString['Peak'] += " kWh"
    # dfTotalsString['Max kVA'] += " kVA"
    # dfTotalsString.to_csv(file, index=False)
    return


def confirmSched(btn):
    count = 0
    for x, row in dfLinks.iterrows():
        dfLinks.loc[x]['State'] = btn[x]['text']
        count = row['Row']
        if row['realColumn'] == 'Sum Week':
            dfF.loc[count] = list([row['Row']]) + list([row['State']]) + ['Nan'] + ['Nan'] + ['Nan'] + ['Nan'] + ['Nan']
        elif row['realColumn'] == 'Sum Sat':
            dfF.loc[count]['Sum Sat'] = row['State']
        elif row['realColumn'] == 'Sum Sun':
            dfF.loc[count]['Sum Sun'] = row['State']
        elif row['realColumn'] == 'Wint Week':
            dfF.loc[count]['Wint Week'] = row['State']
        elif row['realColumn'] == 'Wint Sat':
            dfF.loc[count]['Wint Sat'] = row['State']
        elif row['realColumn'] == 'Wint Sun':
            dfF.loc[count]['Wint Sun'] = row['State']
        # dfF.to_csv('dataFrame.csv', index=False)

    return


# Begin construction of frame
summer = tk.Label(root, text="Summer")
summer.grid(column=2, row=0)
winter = tk.Label(root, text="Winter")
winter.grid(column=9, row=0)
head1 = tk.Label(root, text="Hour")
head1.grid(column=0, row=1)
head2 = tk.Label(root, text="Weekdays")
head2.grid(column=1, row=1)
head3 = tk.Label(root, text="Saturdays")
head3.grid(column=2, row=1)
head4 = tk.Label(root, text="Sunday")
head4.grid(column=3, row=1)
head5 = tk.Label(root, text="      ")
head5.grid(column=4, row=1)
head6 = tk.Label(root, text="      ")
head6.grid(column=5, row=1)
head7 = tk.Label(root, text="      ")
head7.grid(column=6, row=1)
head8 = tk.Label(root, text="      ")
head8.grid(column=7, row=1)
head9 = tk.Label(root, text="Weekdays")
head9.grid(column=8, row=1)
head10 = tk.Label(root, text="Saturdays")
head10.grid(column=9, row=1)
head11 = tk.Label(root, text="Sunday")
head11.grid(column=10, row=1)

lbl0 = tk.Label(root, text="0")
lbl0.grid(column=0, row=2)
lbl1 = tk.Label(root, text="1")
lbl1.grid(column=0, row=3)
lbl2 = tk.Label(root, text="2")
lbl2.grid(column=0, row=4)
lbl3 = tk.Label(root, text="3")
lbl3.grid(column=0, row=5)
lbl4 = tk.Label(root, text="4")
lbl4.grid(column=0, row=6)
lbl5 = tk.Label(root, text="5")
lbl5.grid(column=0, row=7)
lbl6 = tk.Label(root, text="6")
lbl6.grid(column=0, row=8)
lbl7 = tk.Label(root, text="7")
lbl7.grid(column=0, row=9)
lbl8 = tk.Label(root, text="8")
lbl8.grid(column=0, row=10)
lbl9 = tk.Label(root, text="9")
lbl9.grid(column=0, row=11)
lbl10 = tk.Label(root, text="10")
lbl10.grid(column=0, row=12)
lbl11 = tk.Label(root, text="11")
lbl11.grid(column=0, row=13)
lbl12 = tk.Label(root, text="12")
lbl12.grid(column=0, row=14)
lbl13 = tk.Label(root, text="13")
lbl13.grid(column=0, row=15)
lbl14 = tk.Label(root, text="14")
lbl14.grid(column=0, row=16)
lbl15 = tk.Label(root, text="15")
lbl15.grid(column=0, row=17)
lbl16 = tk.Label(root, text="16")
lbl16.grid(column=0, row=18)
lbl17 = tk.Label(root, text="17")
lbl17.grid(column=0, row=19)
lbl18 = tk.Label(root, text="18")
lbl18.grid(column=0, row=20)
lbl19 = tk.Label(root, text="19")
lbl19.grid(column=0, row=21)
lbl20 = tk.Label(root, text="20")
lbl20.grid(column=0, row=22)
lbl21 = tk.Label(root, text="21")
lbl21.grid(column=0, row=23)
lbl22 = tk.Label(root, text="22")
lbl22.grid(column=0, row=24)
lbl23 = tk.Label(root, text="23")
lbl23.grid(column=0, row=25)

fill1 = tk.Label(root, text="      ")
fill1.grid(column=14, row=0)

winterT = tk.Label(root, text="Winter Tariffs")
winterT.grid(column=16, row=0)

fill1 = tk.Label(root, text="      ")
fill1.grid(column=17, row=0)

summerT = tk.Label(root, text="Summer Tariffs")
summerT.grid(column=18, row=0)

peak = tk.Label(root, text="Peak")
peak.grid(column=15, row=1)
off = tk.Label(root, text="Off")
off.grid(column=15, row=2)
standard = tk.Label(root, text="Standard")
standard.grid(column=15, row=3)

winterP = tk.Entry(root)
winterP.grid(column=16, row=1)
winterO = tk.Entry(root)
winterO.grid(column=16, row=2)
winterS = tk.Entry(root)
winterS.grid(column=16, row=3)
summerP = tk.Entry(root)
summerP.grid(column=18, row=1)
summerO = tk.Entry(root)
summerO.grid(column=18, row=2)
summerS = tk.Entry(root)
summerS.grid(column=18, row=3)

stringT = 'O'
stringC = 'Green'
number = 0
for i in range(6):
    for j in range(24):
        if ((j > 7 and j < 13) or (j > 17 and j < 21)) and (i == 0):
            stringT = 'P'
            stringC = 'Red'
        elif ((j == 6) or (j == 7) or (j > 12 and j < 18) or (j == 21)) and (i == 0):
            stringT = 'S'
            stringC = 'Yellow'
        elif ((j > 6 and j < 12) or (j == 18) or (j == 19)) and (i == 1):
            stringT = 'S'
            stringC = 'Yellow'
        elif ((j > 6 and j < 12) or (j > 16 and j < 20)) and (i == 3):
            stringT = 'P'
            stringC = 'Red'
        elif ((j == 5) or (j == 6) or (j > 11 and j < 17) or (j == 20)) and (i == 3):
            stringT = 'S'
            stringC = 'Yellow'
        elif ((j > 5 and j < 11) or (j == 17) or (j == 18)) and (i == 4):
            stringT = 'S'
            stringC = 'Yellow'
        else:
            stringT = 'O'
            stringC = 'Green'
        btn.append(tk.Button(root, text=stringT, padx=20, pady=0, fg="Black", bg=stringC,
                             command=lambda column=i, row=j, nbr=number: changeState(btn, nbr, column, row)))

        if i < 3:
            btn[number].grid(column=i + 1, row=j + 2)
        else:
            btn[number].grid(column=i + 5, row=j + 2)

        realColumn = ''
        if i == 0:
            realColumn = 'Sum Week'
        elif i == 1:
            realColumn = 'Sum Sat'
        elif i == 2:
            realColumn = 'Sum Sun'
        elif i == 3:
            realColumn = 'Wint Week'
        elif i == 4:
            realColumn = 'Wint Sat'
        elif i == 5:
            realColumn = 'Wint Sun'

        dfLinks.loc[number] = [i] + [j] + [number] + list(btn[number]['text']) + [realColumn]
        number += 1

canvas = tk.Canvas(bg="#F0F0F0", width=50)
canvas.grid(column=11, columnspan=5, row=2, rowspan=21)

inputBt = tk.Button(root, text="1. Select Input", bg="Grey", command=selectIn, width=16)
inputBt.grid(column=16, columnspan=3, row=11, rowspan=2)

confirmSchedule = tk.Button(root, text="2. Confirm Schedule", bg="Grey", command=lambda: confirmSched(btn), width=16)
confirmSchedule.grid(column=16, columnspan=3, row=13, rowspan=2)

outputBt = tk.Button(root, text="3. Generate Output", bg="Grey", command=genOutput, width=16)
outputBt.grid(column=16, columnspan=3, row=15, rowspan=2)

quitBt = tk.Button(root, text="4. Exit", bg="Grey", command=exit, width=16)
quitBt.grid(column=16, columnspan=3, row=17, rowspan=2)

testBt = tk.Button(root, text="5. Text", bg="Grey", command=test, width=16)
testBt.grid(column=16, columnspan=3, row=19, rowspan=2)

root.mainloop()
