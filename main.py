""" This script is going to be timecalculating my machiningtime,
    and outputting the results in a excelfile """

#TODO 
# Save the time in seconds instead in the raw data sheet to simplify the results sheet writing

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

time_in_sec = int(round(time_in_sec, 0))

#print(time_in_sec)

seconds = 0
minutes = 0
hours = 0
"""
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

total_time = str(hours) + ':' + str(minutes) + ':' + str(seconds)
"""
if os.path.isfile('Tidskalkyle.xlsx') == False:
    print('Hello')
    wb = op.Workbook()

    wb.create_sheet('Raw Data')
    wb.create_sheet('Results')

    del wb['Sheet']

    ws = wb['Raw Data']
    ws.cell(row=1, column=1, value='Productcode:')
    ws.cell(row=1, column=2, value='Time:')

    ws.cell(row=2, column=1, value=productcode)
    ws.cell(row=2, column=2, value=time_in_sec)

    ws = wb['Results']
    ws.cell(row=1, column=1, value='Productcode:')
    ws.cell(row=1, column=2, value='Total Machining time:')

    wb.save('Tidskalkyle.xlsx')

else:
    wb = op.load_workbook('Tidskalkyle.xlsx')
    ws = wb['Raw Data']

    row = 2
    count = 0
    while ws.cell(row=row, column=1).value != None:

        if ws.cell(row=row, column=1).value == productcode:

            ws.cell(row=row, column=1, value=productcode)
            ws.cell(row=row, column=2, value=time_in_sec)
            count = 1
        
        row += 1
        
    if count == 0:

        ws.cell(row=row, column=1, value=productcode)
        ws.cell(row=row, column=2, value=time_in_sec)

    wb.save('Tidskalkyle.xlsx')

#############################################
# Dealing with the results sheet under      #
#############################################

test = []
row = 2
while ws.cell(row=row, column=1).value != None:
    testvalue = ws.cell(row=row, column=1).value
    test.append(testvalue[:6])
    row += 1

test = list(set(test))
print(test)

ordernumber = ""
time = 0
row = 2
while ws.cell(row=row, column=1).value != None:
    cellvalue = ws.cell(row=row, column=1).value
    cellvalue = cellvalue[:6]

    for item in test:
        if item == cellvalue:
            print('Item:', item, 'Cellvalue:', cellvalue)
            time += ws.cell(row=row, column=2).value
            ordernumber = item
            print('Time', time)

    ws = wb['Results']
    check = 0
    row_results = 2
    while ws.cell(row=row_results, column=1).value != None:
        if ordernumber == ws.cell(row=row_results, column=1).value:
            ws.cell(row=row_results, column=2, value=time)
            check = 1
        row_results += 1
    
    if check == 0:
        ws.cell(row=row_results, column=1, value=ordernumber)
        ws.cell(row=row_results, column=2, value=time)

    ws = wb['Raw Data']
    #time = 0
    row += 1

wb.save('Tidskalkyle.xlsx')

