import openpyxl

# wb = Workbook()
wb = openpyxl.load_workbook('contest-test.xlsx')

# dict = {'ACN/AFS Curriculum': 'CDP'}

#loop through rows and populate category(AA) value
ws = wb.active

for row in ws.iter_rows('AB{}:AB{}'.format(ws.min_row,ws.max_row)):
    value = ws['AB'+str(c)]
    print value

for row in ws.iter_rows('AA{}:AA{}'.format(ws.min_row,ws.max_row)):
    for cell in row:
        # cell.value = 'ACN/AFS Curriculum'
        print cell.value

# Save the file
# wb.save("contest-test.xlsx")

