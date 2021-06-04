import tkinter
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import os
import datetime as DT

main_win = tkinter.Tk()
main_win.title('Attendance Keeper')
main_win.sourceFolder = ''
def chooseDir():
    main_win.sourceFolder = tkinter.filedialog.askdirectory()


def callback():

    global buttonClicked
    buttonClicked = not buttonClicked


buttonClicked = False  # Before first click


tkinter.Label(main_win,
          text="Minimum attendance duration in minutes").grid(row=0)
tkinter.Label(main_win,
         text="Max absence days before disqualification").grid(row=1)
tkinter.Label(main_win,
         text="Lecture start time hh:mm:ss AM/PM").grid(row=2)
tkinter.Label(main_win,
         text="Lecture finish time hh:mm:ss AM/PM").grid(row=3)

e1 = tkinter.Entry(main_win)
e2 = tkinter.Entry(main_win)
e3 = tkinter.Entry(main_win)
e4 = tkinter.Entry(main_win)

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)
e3.grid(row=2, column=1)
e4.grid(row=3, column=1)

tkinter.Button(main_win,
          text='Done',
          command=main_win.quit).grid(row=5,
                                    column=0,
                                    sticky=tkinter.W,
                                    pady=4)

tkinter.Button(main_win,
          text='Browse', command=chooseDir).grid(row=4,
                                                       column=0,
                                                       sticky=tkinter.W,
                                                       pady=4)

main_win.mainloop()

# if start:
# print('Helloooooo\n', main_win.sourceFolder)
ktab = int(e2.get())  # maximum number of absences before 7erman
minmin = int(e1.get())  # minimum duration of minutes, less than that is considered absence
finish = e4.get()
start = e3.get()
xlsx = pd.ExcelFile(os.path.join(main_win.sourceFolder, 'Student List.xlsx'))
file = pd.read_excel(xlsx, 'Sheet1')
reg = file.set_index('Full Name')

# cwd = os.path.abspath('')
files = os.listdir(main_win.sourceFolder)

for each in files:
    if each == 'Student List.xlsx':
        continue
    if each == 'Attendance Sheet.xlsx':
        continue
    if each.endswith('.xlsx'):
        loc = each.find('(')
        loc2 = each.find(')')
        filename = each[loc:loc2+1]
        df = pd.read_excel(each)
        new = pd.merge(file, df, on='Full Name', how='left')
        new.fillna(0, inplace=True)
        record = pd.Series(dtype='float64')
        before = 0
        FMT = '%m/%d/%Y, %H:%M:%S %p'
        for row in new.index:
            try:
                if new['Full Name'][row] == new['Full Name'][row - 1]:
                    continue
            except:
                pass
            if new['Timestamp'][row] == 0:
                ans = pd.Series({new['Full Name'][row]: 0})
                record = record.append(ans)
                continue
            i = 0
            j = 1
            aa = 0
            temp = pd.DataFrame()
            reg = new['Full Name'][row]
            new = new.set_index('Full Name', drop=False)
            temp = temp.append(new.loc[reg])
            new = new.reset_index(drop=True)
            before = len(temp.index)
            time = new['Timestamp'][row]
            end = time[:12] + finish
            begin = time[:12] + start
            k = 0
            if before == 1:
                if DT.datetime.strptime(time, FMT) < DT.datetime.strptime(begin, FMT):
                    time = begin
                val = DT.datetime.strptime(end, FMT) - DT.datetime.strptime(time, FMT)
                val = round(val.total_seconds() / 60, 1)
                if val > 75: val = 75
                ans = pd.Series({new['Full Name'][row]: val})
                record = record.append(ans)
                continue
            if before % 2 == 0:
                net = before // 2
                while aa < net:
                    tbfr = new['Timestamp'][row + i]
                    taft = new['Timestamp'][row + j]
                    if DT.datetime.strptime(tbfr, FMT) < DT.datetime.strptime(begin, FMT):
                        tbfr = begin
                    if DT.datetime.strptime(taft, FMT) < DT.datetime.strptime(begin, FMT):
                        taft = begin
                    k = DT.datetime.strptime(taft, FMT) - DT.datetime.strptime(tbfr, FMT) + DT.timedelta(k)
                    i += 2
                    j += 2
                    aa += 1
                    k = round(k.total_seconds() / 60, 1)
                if k > 75: k = 75
                ans = pd.Series({new['Full Name'][row]: k})
                record = record.append(ans)
                continue
            if before % 2 != 0:
                net = before // 2
                last = new['Timestamp'][row + before - 1]
                if DT.datetime.strptime(last, FMT) < DT.datetime.strptime(begin, FMT):
                    last = begin
                k = DT.datetime.strptime(end, FMT) - DT.datetime.strptime(last, FMT) + DT.timedelta(k)
                while aa < net:
                    k = k.total_seconds()
                    tbfr = new['Timestamp'][row + i]
                    taft = new['Timestamp'][row + j]
                    if DT.datetime.strptime(tbfr, FMT) < DT.datetime.strptime(begin, FMT):
                        tbfr = begin
                    if DT.datetime.strptime(taft, FMT) < DT.datetime.strptime(begin, FMT):
                        taft = begin
                    k = DT.datetime.strptime(taft, FMT) - DT.datetime.strptime(tbfr, FMT) + DT.timedelta(seconds=k)
                    i += 2
                    j += 2
                    aa += 1
                k = round(k.total_seconds() / 60, 1)
                if k > 75: k = 75
                ans = pd.Series({new['Full Name'][row]: k})
                record = record.append(ans)
                continue
        record.index.name = 'Full Name'
        record = record.to_frame(name=filename)
        file = pd.merge(file, record, on='Full Name', how='left')
file.set_index('Full Name', inplace=True)
file.drop('N', axis='columns', inplace=True)
n = file.values < minmin
ab = n.sum(axis=1)
file['Absence'] = ab
herman = pd.Series(dtype='str')
for row in file['Absence']:
    if row > ktab:
        dd = pd.Series({file['Absence'][row]: 'Yes'})
        herman = herman.append(dd)
    else:
        dd = pd.Series({file['Absence'][row]: 'No'})
        herman = herman.append(dd)

file['ktab'] = herman.values
print(file)
writer = pd.ExcelWriter(main_win.sourceFolder, engine='xlsxwriter')
file.to_excel(writer, 'Attendance Sheet.xlsx')
writer.save()

# 10:00:00 AM
# 11:15:00 AM