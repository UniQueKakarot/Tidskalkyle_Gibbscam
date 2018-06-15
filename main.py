""" This script is going to be timecalculating my machiningtime,
    and outputting the results in a excelfile """

import os.path

import openpyxl as op

with open('Time4.txt') as f:
    filebody = [line for line in f]

filebody_clean = [item[:-1] for item in filebody]
productcode = filebody_clean.pop(0)

filebody_clean = [float(item) for item in filebody_clean]

time_in_sec = 0.0
for time in filebody_clean:
    time_in_sec += time

time_in_sec = round(time_in_sec, 0)

print(time_in_sec)

seconds = 0
minutes = 0
hours = 0

while time_in_sec != 0.0:

    if time_in_sec > 3600.0:
        hours += 1
        time_in_sec -= 3600.0

    elif time_in_sec > 60.0:
        minutes += 1
        time_in_sec -= 60.0

    else:
        seconds = time_in_sec
        time_in_sec -= time_in_sec

print('Hours: ', hours)
print('Minutes: ', minutes)
print('Seconds: ', seconds)
print(time_in_sec)

total_time = str(hours) + ':' + str(minutes) + ':' + str(seconds)

print(total_time)

if os.path.isfile('Tidskalkyle.xlsx') == False:
    wb = op.Workbook()

    wb.create_sheet('Raw Data')
    wb.create_sheet('Results')

    del wb['Sheet']

    ws = wb['Raw Data']
    ws.cell(row=1, column=1, value='Productcode:')
    ws.cell(row=1, column=2, value='Time:')

    ws.cell(row=2, column=1, value=productcode)
    ws.cell(row=2, column=2, value=total_time)

    ws = wb['Results']
    ws.cell(row=1, column=1, value='Productcode:')
    ws.cell(row=1, column=2, value='Total Machining time:')

    raw_productcode = productcode[0:6]
    ws.cell(row=2, column=1, value=raw_productcode)
    ws.cell(row=2, column=2, value=total_time)

    wb.save('Tidskalkyle.xlsx')

else:
    wb = op.load_workbook('Tidskalkyle.xlsx')
