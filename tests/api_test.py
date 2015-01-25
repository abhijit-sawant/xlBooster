#!usr/bin/env python

import os
import win32com
import xlbooster.constants as constants
import xlbooster.xlb       as xlb

app = xlb.xlbApp()
strPathWb = 'C:\\Users\\asawant\\PlayGround\\GitHub\\xlBooster\\tests\\test'
strNameWs = 'Data'

if os.path.exists(strPathWb + '.xlsx'):
    os.remove(strPathWb + '.xlsx')

#-------------------------------------------------------------------------------
def setVals():
    wb = app.addWorkBook()
    ws = wb.addWorkSheet()
    ws.setName(strNameWs)
    r = ws.getRange(1, 1, 10, 2)

    l = [[i+j for i in range(2)] for j in range(10)]
    r.setVals(l)

    wb.saveAs(strPathWb)
    wb.close()

#-------------------------------------------------------------------------------
def getVals():
    wb = app.openWorkBook(strPathWb)
    ws = wb.getWorkSheet(strNameWs)
    r = ws.getRange(1, 1, 10, 2)

    l = r.getVals()
    print l

#-------------------------------------------------------------------------------
def setVis():
    wb = app.openWorkBook(strPathWb)
    ws = wb.getWorkSheet(strNameWs)
    r = ws.getRange(1, 1, 10, 2)

    r.setFillColor(11919826)
    r.setFontColor(38400)
    r.setFont('arial', 'bold')
    r.setBorder(constants.xlMedium)

#-------------------------------------------------------------------------------
def setArray():
    import numpy as np

    wb = app.addWorkBook()
    ws = wb.getWorkSheet('Sheet1')
    r = ws.getRange(1, 1, 2, 2)

    ar = np.array([[1, 11], [2, 22]])
    r.setArray(ar)

#-------------------------------------------------------------------------------
def getArray():
    wb = app.getWorkBook('Book1')
    ws = wb.getWorkSheet('Sheet1')
    r = ws.getRange(1, 1, 2, 2)
    
    ar = r.getArray()
    print ar

#-------------------------------------------------------------------------------
def createBarChartStacked():
    chart = ws.addChart(r, constants.xlColumnStacked)
    #chart.setName('bar_chart')
    #bar_chart = ws.getChart('bar_chart')

#End of file
