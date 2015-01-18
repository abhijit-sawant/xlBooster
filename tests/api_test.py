#!usr/bin/env python

import win32com
import xlb

app = xlb.xlbApp()

#-------------------------------------------------------------------------------
def setList():
    wb = app.addWorkBook()
    ws = wb.addWorkSheet()

    r = ws.getRange(1, 1, 10, 2)

    l = [[i+j for i in range(2)] for j in range(10)]
    r.setList(l)

#-------------------------------------------------------------------------------
def setArray():
    import numpy as np

    wb = app.addWorkBook()
    ws = wb.getWorkSheet('Sheet1')

    r = ws.getRange(1, 1, 2, 2)

    ar = np.array([[1, 11], [2, 22]])
    r.setArray(ar)

#-------------------------------------------------------------------------------
def createBarChartStacked():
    chart = ws.addChart(r, win32com.client.constants.xlColumnStacked)
    #chart.setName('bar_chart')
    #bar_chart = ws.getChart('bar_chart')

#End of file
