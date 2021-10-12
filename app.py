## XLSX TO CSV
import openpyxl

filename = 'Sample_source.xlsm'

## opening the xlsx file
xlsx = openpyxl.load_workbook(filename)

## opening the active sheet
sheet = xlsx.active

## getting the data from the sheet
data = sheet.rows

## creating a csv file
csv = open("data.csv", "w+")

count=0

for row in data:

    print(row)
    if(count>=3):
        csv.write('\n')
        l = list(row)
        for i in range(len(l)):
            if i == len(l) - 1:
                csv.write(str(l[i].value))
            else:
                csv.write(str(l[i].value) + ',')
    count=count+1

## close the csv file
csv.close()