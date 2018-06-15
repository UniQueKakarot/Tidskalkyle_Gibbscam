""" This script is going to be timecalculating my machiningtime,
    and outputting the results in a excelfile """

with open('Time4.txt') as f:
    filebody = [line for line in f]

filebody_clean = [item[:-1] for item in filebody]
filename = filebody_clean.pop(0)

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
