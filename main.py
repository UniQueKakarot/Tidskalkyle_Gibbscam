""" This script is going to be timecalculating my machiningtime,
    and outputting the results in a excelfile """

with open('Time.txt') as f:
    filebody = [line for line in f]

filebody_clean = [item[:-1] for item in filebody]
filename = filebody_clean.pop(0)

filebody_clean = [float(item) for item in filebody_clean]

time_in_sec = 0.0
for time in filebody_clean:
    time_in_sec += time 

print(time_in_sec)

seconds = 0
minutes = 0
hours = 0


