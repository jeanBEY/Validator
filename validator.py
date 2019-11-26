#!/usr/bin/env python3
#validator.py - Validates the data on STRS/PERS report for submission to LACOE

import openpyxl, pprint
from openpyxl.styles import Font, Color, PatternFill, Border
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.comments import Comment

print('Opening workbook...')
wb = openpyxl.load_workbook('LA Co Of Ed.xlsx')
sheet = wb.active

end = 1

#Find last row with data in column A
for lastRow in range(2, sheet.max_row):
    if not sheet['A' + str(lastRow)].value:
        end = lastRow
        break

#Color blank cells black & input text "BLANK"
for row in range(2, end):
    for col in range(1, column_index_from_string('AB')):
        if not sheet[get_column_letter(col) + str(row)].value:
            sheet[get_column_letter(col) + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
            sheet[get_column_letter(col) + str(row)].value = 'BLANK'

#Agency name (column A) should be the same value.  Uses value in A2 as correct #.
for row in range(3, end):
    if not (sheet['A' + str(2)].value == sheet['A' + str(row)].value):
        comment = Comment("Agency # does not match", "Windows User")
        sheet['A' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['A' + str(row)].comment = comment

#Gender must be either M or F
for row in range(2, end):    
    if not (sheet['D' + str(row)].value == 'M' or sheet['D' + str(row)].value == 'F'):
        comment = Comment("Gender must be M or F", "Windows User")
        sheet['D' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['D' + str(row)].comment = comment

#Classification must be 100010 or 200010
for row in range(2, end):    
    if not (sheet['H' + str(row)].value == 100010 or sheet['H' + str(row)].value == 200010):
        comment = Comment("Classification must be 100010 or 200010", "Windows User")
        sheet['H' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['H' + str(row)].comment = comment

#Transaction Code must be either TX, LX, RA or RX
for row in range(2, end):    
    if not (sheet['M' + str(row)].value == 'TX' or sheet['M' + str(row)].value == 'LX' or sheet['M' + str(row)].value == 'RX' or sheet['M' + str(row)].value == 'RA'):
        comment = Comment("Transaction code must be TX/LX/RA/RX", "Windows User")
        sheet['M' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['M' + str(row)].comment = comment

#Earning Type must be either REG, OVT, RTS
for row in range(2, end):    
    if not (sheet['N' + str(row)].value == 'REG' or sheet['N' + str(row)].value == 'OVT' or sheet['N' + str(row)].value == 'RTS'):
        comment = Comment("Earning type must be REG/OVT/RTS", "Windows User")
        sheet['N' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['N' + str(row)].comment = comment
        
wb.save('LA Co Of Ed - UPDATED.xlsx')
