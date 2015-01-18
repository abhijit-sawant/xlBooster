#!usr/bin/env python

import win32com
import xlb

app = xlb.xlbApp()

wb = app.addWorkBook()
ws = wb.addWorkSheet()
ws.activate()

r = ws.getRange(1, 1, 10, 2)

l = [[i+j for i in range(2)] for j in range(10)]

r.setList(l)

chart = ws.addChart(r, win32com.client.constants.xlColumnStacked)

#End of file
