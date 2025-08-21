# Week 2: Local Excel parser script for RAID.xlsx
import openpyxl

# Load local RAID.xlsx file
wb = openpyxl.load_workbook("RAID_fixed.xlsx")
sheet = wb.active

# Print all rows
for row in sheet.iter_rows(values_only=True):
    print(row)

