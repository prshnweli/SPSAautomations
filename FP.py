#imports
import pandas as pd
import pyodbc
import numpy as np
import csv
import openpyxl
import pygsheets
import config
gc = pygsheets.authorize(service_file='keys/creds.json')

# Connect to database
conn = pyodbc.connect('Driver={SQL Server};'
'Server='+config.Server+';'
'Database='+config.Database+';'
'UID='+config.UID+';'
'PWD='+config.PWD+';'
)

# SQL File
file = open('sql/fp.sql', 'r')
sqlFile = file.read()
file.close()

def xl(sc, output):
    # Run SQL query with connection and store in dataframe.
    writer = pd.ExcelWriter(output,engine='xlsxwriter')
    workbook=writer.book
    worksheet=workbook.add_worksheet('FP')
    writer.sheets['FP'] = worksheet

    bl = pd.read_sql_query("DECLARE @SchoolNum AS INT =" + sc + "DECLARE @Gbknum AS INT = 900 " + sqlFile, conn)
    t2 = pd.read_sql_query("DECLARE @SchoolNum AS INT =" + sc + "DECLARE @Gbknum AS INT = 902 " + sqlFile, conn)
    # Write to xlsx
    bl.to_excel(writer,sheet_name='FP',startrow=1 , startcol=8, index = False) #Puts baseline dataframe at I2
    #
    t2.to_excel(writer,sheet_name='FP',startrow=1 , startcol=17, index = False) #Puts trimester 2 dataframe at R2


    # Write headers for xlsx
    worksheet.write_string('A1', "Fall to Winter Growth by Grade Level")
    worksheet.write_string('B2', "0")
    worksheet.write_string('C2', "1")
    worksheet.write_string('D2', "2")
    worksheet.write_string('E2', "3")
    worksheet.write_string('F2', "4")
    worksheet.write_string('G2', "5")

    worksheet.write_string('I1', "Fall F&P")
    worksheet.write_string('R1', "Winter F&P")

    # Column details
    worksheet.write_string('A2', "Grade")
    worksheet.write_string('A3', "Fall Avg")
    worksheet.write_string('A4', "Winter Avg")
    worksheet.write_string('A5', "Growth")
    worksheet.write_string('A6', "% At Level")

    # Formulas
    worksheet.write_formula('C3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 1, O3:O999, ">0"),2)') # Grade 1 FALL
    worksheet.write_formula('D3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 2, O3:O999, ">0"),2)') # Grade 2 FALL
    worksheet.write_formula('E3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 3, O3:O999, ">0"),2)') # Grade 3 FALL
    worksheet.write_formula('F3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 4, O3:O999, ">0"),2)') # Grade 4 FALL
    worksheet.write_formula('G3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 5, O3:O999, ">0"),2)') # Grade 5 FALL

    worksheet.write_formula('B4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 0, X3:W999, ">0"),2)') # Grade 0 WINTER
    worksheet.write_formula('C4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 1, X3:W999, ">0"),2)') # Grade 1 WINTER
    worksheet.write_formula('D4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 2, X3:W999, ">0"),2)') # Grade 2 WINTER
    worksheet.write_formula('E4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 3, X3:W999, ">0"),2)') # Grade 3 WINTER
    worksheet.write_formula('F4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 4, X3:W999, ">0"),2)') # Grade 4 WINTER
    worksheet.write_formula('G4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 5, X3:W999, ">0"),2)') # Grade 5 WINTER

    worksheet.write_formula('B5', '=B4-B3') # Grade 0 Growth
    worksheet.write_formula('C5', '=C4-C3') # Grade 1 Growth
    worksheet.write_formula('D5', '=D4-D3') # Grade 2 Growth
    worksheet.write_formula('E5', '=E4-E3') # Grade 3 Growth
    worksheet.write_formula('F5', '=F4-F3') # Grade 4 Growth
    worksheet.write_formula('G5', '=G4-G3') # Grade 5 Growth

    worksheet.write_formula('B6', '=ROUND(COUNTIFS(U3:T999, 0, X3:W999, ">=0")/COUNTIF(U3:T999, 0),4)') # Grade 0 At Grade Level
    worksheet.write_formula('C6', '=ROUND(COUNTIFS(U3:T999, 1, X3:W999, ">=1")/COUNTIF(U3:T999, 1),4)') # Grade 0 At Grade Level
    worksheet.write_formula('D6', '=ROUND(COUNTIFS(U3:T999, 2, X3:W999, ">=2")/COUNTIF(U3:T999, 2),4)') # Grade 0 At Grade Level
    worksheet.write_formula('E6', '=ROUND(COUNTIFS(U3:T999, 3, X3:W999, ">=3")/COUNTIF(U3:T999, 3),4)') # Grade 0 At Grade Level
    worksheet.write_formula('F6', '=ROUND(COUNTIFS(U3:T999, 4, X3:W999, ">=4")/COUNTIF(U3:T999, 4),4)') # Grade 0 At Grade Level
    worksheet.write_formula('G6', '=ROUND(COUNTIFS(U3:T999, 5, X3:W999, ">=5")/COUNTIF(U3:T999, 5),4)') # Grade 0 At Grade Level

    writer.save()

def goog(sc, sheet):
    bl = pd.read_sql_query("DECLARE @SchoolNum AS INT =" + sc + "DECLARE @Gbknum AS INT = 900 " + sqlFile, conn)

    t2 = pd.read_sql_query("DECLARE @SchoolNum AS INT =" + sc + "DECLARE @Gbknum AS INT = 902 " + sqlFile, conn)

    df = pd.DataFrame(bl)
    df2 = pd.DataFrame(t2)

    sh = gc.open(sheet)
    wks = sh[0]
    wks.clear()
    wks.set_dataframe(df,(2,9))
    wks.set_dataframe(df2,(2,18))

    wks.update_value('A1', "Fall to Winter Growth by Grade Level")
    wks.update_value('B2', "0")
    wks.update_value('C2', "1")
    wks.update_value('D2', "2")
    wks.update_value('E2', "3")
    wks.update_value('F2', "4")
    wks.update_value('G2', "5")

    wks.update_value('I1', "Fall F&P")
    wks.update_value('R1', "Winter F&P")

    # Column details
    wks.update_value('A2', "Grade")
    wks.update_value('A3', "Fall Avg")
    wks.update_value('A4', "Winter Avg")
    wks.update_value('A5', "Growth")
    wks.update_value('A6', "% At Level")

    # Formulas
    wks.update_value('C3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 1, O3:O999, ">0"),2)') # Grade 1 FALL
    wks.update_value('D3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 2, O3:O999, ">0"),2)') # Grade 2 FALL
    wks.update_value('E3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 3, O3:O999, ">0"),2)') # Grade 3 FALL
    wks.update_value('F3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 4, O3:O999, ">0"),2)') # Grade 4 FALL
    wks.update_value('G3', '=ROUND(AVERAGEIFS(O3:O999, L3:L999, 5, O3:O999, ">0"),2)') # Grade 5 FALL

    wks.update_value('B4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 0, X3:W999, ">0"),2)') # Grade 0 WINTER
    wks.update_value('C4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 1, X3:W999, ">0"),2)') # Grade 1 WINTER
    wks.update_value('D4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 2, X3:W999, ">0"),2)') # Grade 2 WINTER
    wks.update_value('E4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 3, X3:W999, ">0"),2)') # Grade 3 WINTER
    wks.update_value('F4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 4, X3:W999, ">0"),2)') # Grade 4 WINTER
    wks.update_value('G4', '=ROUND(AVERAGEIFS(X3:W999, U3:T999, 5, X3:W999, ">0"),2)') # Grade 5 WINTER

    wks.update_value('B5', '=B4-B3') # Grade 0 Growth
    wks.update_value('C5', '=C4-C3') # Grade 1 Growth
    wks.update_value('D5', '=D4-D3') # Grade 2 Growth
    wks.update_value('E5', '=E4-E3') # Grade 3 Growth
    wks.update_value('F5', '=F4-F3') # Grade 4 Growth
    wks.update_value('G5', '=G4-G3') # Grade 5 Growth

    wks.update_value('B6', '=ROUND(COUNTIFS(U3:T999, 0, X3:W999, ">=0")/COUNTIF(U3:T999, 0),4)') # Grade 0 At Grade Level
    wks.update_value('C6', '=ROUND(COUNTIFS(U3:T999, 1, X3:W999, ">=1")/COUNTIF(U3:T999, 1),4)') # Grade 0 At Grade Level
    wks.update_value('D6', '=ROUND(COUNTIFS(U3:T999, 2, X3:W999, ">=2")/COUNTIF(U3:T999, 2),4)') # Grade 0 At Grade Level
    wks.update_value('E6', '=ROUND(COUNTIFS(U3:T999, 3, X3:W999, ">=3")/COUNTIF(U3:T999, 3),4)') # Grade 0 At Grade Level
    wks.update_value('F6', '=ROUND(COUNTIFS(U3:T999, 4, X3:W999, ">=4")/COUNTIF(U3:T999, 4),4)') # Grade 0 At Grade Level
    wks.update_value('G6', '=ROUND(COUNTIFS(U3:T999, 5, X3:W999, ">=5")/COUNTIF(U3:T999, 5),4)') # Grade 0 At Grade Level



xl("2", 'FP/ElToro.xlsx')
xl("6", 'FP/LosPaseos.xlsx')
xl("8", 'FP/Nordstrom.xlsx')
xl("9", 'FP/Paradise.xlsx')
xl("10", 'FP/SMG.xlsx')
xl("11", 'FP/Walsh.xlsx')
xl("12", 'FP/Barrett.xlsx')
xl("15", 'FP/JAMM.xlsx')

goog("2",'El Toro FP')
goog("6", 'Los Paseos FP')
goog("8", 'Nordstrom FP')
goog("9", 'Paradise FP')
goog("10", 'SMG FP')
goog("11", 'Walsh FP')
goog("12", 'Barrett FP')
goog("15", 'JAMM FP')
