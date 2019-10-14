from openpyxl import load_workbook
from openpyxl import Workbook

wb1 = load_workbook(filename="Book1.xlsx")
sheet = wb1.worksheets[0]
#print(sheet['A8'].value)
data = []
count = 0
# Make a list of lists
# Each list has that year's data
for row in sheet.iter_rows(min_row=1, min_col=1, max_col=1):
    for cell in row:
        if "$ POPULATION BY SINGLE YEAR OF AGE" in str(cell.value):
            data.append([])
            #print(cell)
            # Get 21 rows underneath the title
            end = 'A' + str(int(str(cell.coordinate)[1:]) + 46)
            #data.append(sheet[cell.coordinate: end])
            #sheet2.append(val[0] for val in sheet[cell.coordinate: end])
            for val in sheet[cell.coordinate: end]:
                for obj in val:
                    data[count].append(obj.value)
            count += 1
#print(data)

# Some data is stuck in elements with other data
# Split into two lists, then merge
new_data = []
for year in data:
    part_one = []
    part_two = []
    part_one.append(year[:1])
    for vals in year[2:]:
        #print(vals)
        space = vals.split(' ')
        #print(space)
        if len(space) > 4:
            part_one.append(space[:4])
            part_two.append(space[4:])
        else:
            part_one.append(space[:4])
    new_data.append(part_one + part_two)

fhand = open("results.csv", "w")
for year in new_data:
    fhand.write(''.join(year[0]) + '\n')
    for entry in year[1:]:
        fhand.write(','.join(entry) + '\n')
fhand.close()
# Paste results in new workbook
"""
wb2 = Workbook()
wb2.save("mortpak_output.xlsx")
sheet2 = wb2.worksheets[0]
# for rows in range A1:DXXX
# find DXXX
# set A to
# set B to ...
total  = sum([len(x) for x in new_data])
print(new_data)
"""
