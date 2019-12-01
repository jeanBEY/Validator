#!/usr/bin/env python3
#validator.py - Validates the data on STRS/PERS report for submission to LACOE

import openpyxl, pprint
from openpyxl.styles import Font, Color, PatternFill, Border
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.comments import Comment

from openpyxl.styles.colors import BLUE

print('Opening workbook...')
wb = openpyxl.load_workbook('LA Co Of Ed.xlsx', data_only=True)
sheet = wb.active

end = 1
currentStart = 191001
currentEnd = 191031

#########################
#   DATA VALIDATION
#########################

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

#PEPRA Code must be either * or 1
for row in range(2, end):    
    if not (sheet['V' + str(row)].value == '*' or sheet['V' + str(row)].value == 1):
        comment = Comment("PEPRA Code must be */1", "Windows User")
        sheet['V' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['V' + str(row)].comment = comment

#Retirement Plan must be either S3, S5 or P9
for row in range(2, end):    
    if not (sheet['T' + str(row)].value == 'S3' or sheet['T' + str(row)].value == 'S5' or sheet['T' + str(row)].value == 'P9'):
        comment = Comment("Retirement Plan must be S3/S5/P9", "Windows User")
        sheet['T' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['T' + str(row)].comment = comment

#Retirement Status must be either M, N or R
for row in range(2, end):    
    if not (sheet['U' + str(row)].value == 'M' or sheet['U' + str(row)].value == 'N' or sheet['U' + str(row)].value == 'R'):
        comment = Comment("Retirement Status must be M/N/R", "Windows User")
        sheet['U' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['U' + str(row)].comment = comment

#Rate must be > 0
for row in range(2, end):    
    if not (sheet['R' + str(row)].value > 0):
        comment = Comment("Rate must be higher than zero", "Windows User")
        sheet['R' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['R' + str(row)].comment = comment

#Reporting Rate must be > 0
for row in range(2, end):    
    if not (sheet['AA' + str(row)].value > 0):
        comment = Comment("Reporting rate must be higher than zero", "Windows User")
        sheet['AA' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['AA' + str(row)].comment = comment

#Percentage must be > 0 && <= 100
for row in range(2, end):    
    if not (sheet['Z' + str(row)].value > 0 and sheet['Z' + str(row)].value <= 100):
        comment = Comment("Percentage higher than zero & less than or equal to 100", "Windows User")
        sheet['Z' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['Z' + str(row)].comment = comment

#Session must be S
for row in range(2, end):    
    if not (sheet['AB' + str(row)].value == 'S'):
        comment = Comment("Session must be S", "Windows User")
        sheet['AB' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['AB' + str(row)].comment = comment

#Number of pays must be 10, 11 or 12
for row in range(2, end):    
    if not (sheet['P' + str(row)].value == 10 or sheet['P' + str(row)].value == 11 or sheet['P' + str(row)].value == 12):
        comment = Comment("Number of pays must be 10, 11 or 12", "Windows User")
        sheet['P' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['P' + str(row)].comment = comment

#Time must be > 0
for row in range(2, end):    
    if not (sheet['Q' + str(row)].value > 0):
        comment = Comment("Time must be higher than zero", "Windows User")
        sheet['Q' + str(row)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        sheet['Q' + str(row)].comment = comment

#########################
#        ALL
#########################

#Check current earnings
for row in range(2, end): 
    if (sheet['M' + str(row)].value == 'TX'):
        if not (sheet['I' + str(row)].value >= currentStart and sheet['I' + str(row)].value <= currentEnd):
            comment = Comment("Start date is not current", "Windows User")
            sheet['I' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['I' + str(row)].comment = comment

        if not(sheet['J' + str(row)].value >= currentStart and sheet['J' + str(row)].value <= currentEnd):
            comment = Comment("End date is not current", "Windows User")
            sheet['J' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['J' + str(row)].comment = comment

        if not (sheet['S' + str(row)].value > 0):
            comment = Comment("Earnings must be positive", "Windows User")
            sheet['S' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['S' + str(row)].comment = comment


        if not (sheet['AA' + str(row)].value > 0):
            comment = Comment("Reporting rate must be positive", "Windows User")
            sheet['AA' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['AA' + str(row)].comment = comment

#Check reversal earnings
for row in range(2, end): 
    if (sheet['M' + str(row)].value == 'RX'):
        if not (sheet['I' + str(row)].value < currentStart and sheet['I' + str(row)].value <= sheet['J' + str(row)].value):
            comment = Comment("Start date is not prior", "Windows User")
            sheet['I' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['I' + str(row)].comment = comment

        if not(sheet['J' + str(row)].value < currentStart and sheet['J' + str(row)].value >= sheet['I' + str(row)].value):
            comment = Comment("End date is not prior", "Windows User")
            sheet['J' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['J' + str(row)].comment = comment

        if not (sheet['S' + str(row)].value < 0):
            comment = Comment("Earnings must be negative (reversal)", "Windows User")
            sheet['S' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['S' + str(row)].comment = comment


        if not (sheet['AA' + str(row)].value > 0):
            comment = Comment("Reporting rate must be positive", "Windows User")
            sheet['AA' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['AA' + str(row)].comment = comment

#Check late earnings
for row in range(2, end): 
    if (sheet['M' + str(row)].value == 'LX'):
        if not (sheet['I' + str(row)].value < currentStart and sheet['I' + str(row)].value <= sheet['J' + str(row)].value):
            comment = Comment("Start date is not prior", "Windows User")
            sheet['I' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['I' + str(row)].comment = comment

        if not(sheet['J' + str(row)].value < currentStart and sheet['J' + str(row)].value >= sheet['I' + str(row)].value):
            comment = Comment("End date is not prior", "Windows User")
            sheet['J' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['J' + str(row)].comment = comment

        if not (sheet['S' + str(row)].value > 0):
            comment = Comment("Earnings must be positive", "Windows User")
            sheet['S' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['S' + str(row)].comment = comment


        if not (sheet['AA' + str(row)].value > 0):
            comment = Comment("Reporting rate must be positive", "Windows User")
            sheet['AA' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['AA' + str(row)].comment = comment

#Check stipends
for row in range(2, end): 
    if (sheet['O' + str(row)].value == 'L'):
        if not (sheet['R' + str(row)].value == sheet['S' + str(row)].value and sheet['S' + str(row)].value == sheet['AA' + str(row)].value):
            comment = Comment("This is a stipend.  Rate = earnings = reporting rate.", "Windows User")
            sheet['S' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
            sheet['S' + str(row)].comment = comment

#########################
#      S T R SYSTEM
#########################

#Check hourly earnings
for row in range(2, end):
    #If a STRS line
    if (sheet['H' + str(row)].value == 100010):
        #If an hourly line
        if (sheet['O' + str(row)].value == 'H'):
            
            if not (sheet['Q' + str(row)].value == (sheet['S' + str(row)].value/sheet['R' + str(row)].value)):
                comment = Comment("Time must be earnings/rate & positive", "Windows User")
                sheet['Q' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
                sheet['Q' + str(row)].comment = comment
                
            if not (sheet['R' + str(row)].value < 100):
                comment = Comment("Rate must be an hourly rate, ideally less than $100/hr", "Windows User")
                sheet['R' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
                sheet['R' + str(row)].comment = comment
                
            if not (sheet['AA' + str(row)].value > 25000):
                comment = Comment("Reporting rate must be positive & an annualized rate, ideally over 25k", "Windows User")
                sheet['AA' + str(row)].fill = PatternFill(fgColor=BLUE, fill_type = "solid")
                sheet['AA' + str(row)].comment = comment

                
wb.save('LA Co Of Ed - UPDATED.xlsx')
