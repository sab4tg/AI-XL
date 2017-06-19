from openpyxl import load_workbook

wb2 = load_workbook('contest-train.xlsx')

# Save the file
# wb.save("contest-train.xlsx")

print wb2.get_sheet_names()