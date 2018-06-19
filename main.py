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

print(time_in_sec)

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

order_time = 0
for cell in ws['A']:
    print('Hello')
    try:
        raw1 = cell.value
        raw1 = raw1[0:6]
    except TypeError:
        pass

    row = 1
    for value in ws['A']:
        print('Hello1')
        try:
            raw2 = value.value
            raw2 = raw2[0:6]
        except TypeError:
            pass

        if raw1 == 'Produc':
            print(raw1)
            pass

        elif raw1 == None:
            print(raw1)
            pass

        elif raw1 == raw2:
            time = int(ws.cell(row=row, column=2).value)
            order_time += time
            print('added up some time?')

            ws = wb['Results']
            print('Hello???')
            row1 = 2
            while ws.cell(row=row1, column=1).value != None:
                print('While...')
                if raw1 == 'Produc':
                    pass
                elif raw1 == None:
                    pass
                else:
                    ws.cell(row=row1, column=1, value=raw1)
                    ws.cell(row=row1, column=2, value=order_time)
                row1 += 1

            ws = wb['Raw Data']

        row += 1

wb.save('Tidskalkyle.xlsx')

            
    #print(cell.value)


"""
row = 2
order_time = 0
while ws.cell(row=row, column=1).value != None:
    productcode = ws.cell(row=row, column=1).value
    raw_productcode = productcode[0:6]
    print('Raw_product', raw_productcode)

    row2 = row + 1
    while ws.cell(row=row2, column=1).value != None:
        inner_productcode = ws.cell(row=row2, column=1).value
        raw_inner = inner_productcode[0:6]
        print('raw_inner', raw_inner)

        if raw_inner == raw_productcode:
            print('Hello')
            time = ws.cell(row=row2, column=2).value

        row2 += 1
        #print('Row2', row2)

    row += 1
    order_time
"""

ws = wb['Results']

row = 2
#while ws.cell(row=row, column=1).value != None:
    #pass
