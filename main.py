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

total_time = str(hours) + ':' + str(minutes) + ':' + str(seconds)

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
    ws.cell(row=2, column=2, value=total_time)

    ws = wb['Results']
    ws.cell(row=1, column=1, value='Productcode:')
    ws.cell(row=1, column=2, value='Total Machining time:')

    #raw_productcode = productcode[0:6]
    #ws.cell(row=2, column=1, value=raw_productcode)
    #ws.cell(row=2, column=2, value=total_time)

    wb.save('Tidskalkyle.xlsx')

else:
    print('Hello 1')
    wb = op.load_workbook('Tidskalkyle.xlsx')
    ws = wb['Raw Data']

    row = 2
    count = 0
    while ws.cell(row=row, column=1).value != None:

        if ws.cell(row=row, column=1).value == productcode:
            print('Hello 2')

            ws.cell(row=row, column=1, value=productcode)
            ws.cell(row=row, column=2, value=total_time)
            count = 1
        
        row += 1
        
    if count == 0:
        print('Hello 3')
        ws.cell(row=row, column=1, value=productcode)
        ws.cell(row=row, column=2, value=total_time)

    wb.save('Tidskalkyle.xlsx')

row = 2
order_time = 0
while ws.cell(row=row, column=1).value != None:
    # find every entry with the same production number and add togheter the times
    productcode = ws.cell(row=row, column=1).value
    raw_productcode = productcode[0:6]

    row2 = 2
    while ws.cell(row=row2, column=1).value != None:
        inner_productcode = ws.cell(row=row, column=1).value
        raw_inner = inner_productcode[0:6]
        if raw_inner == raw_productcode:
            time = ws.cell(row=row2, column=2).value


    order_time

ws = wb['Results']

row = 2
while ws.cell(row=row, column=1).value != None:
    pass
