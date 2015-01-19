#!/usr/bin/env python

try:
    from win32com.client import gencache
    gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 6)
except:
    raise Exception('Could not generate required Excel constatns. Import of module failed.')

import math
import types
import win32com.client

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbApp(object):
    #-------------------------------------------------------------------------------
    def __init__(self):
        self.__app = win32com.client.DispatchEx("Excel.Application")
        self.__app.Visible = 1

    #-------------------------------------------------------------------------------
    def addWorkBook(self):
        wb = self.__app.Workbooks.Add()
        return xlbWorkBook(self.__app, wb)

    #-------------------------------------------------------------------------------
    def openWorkBook(self, strPath):
        wb = self.__app.Workbooks.Open(strPath)
        return xlbWorkBook(self.__app, wb)

    #-------------------------------------------------------------------------------
    def getWorkBook(self, strName):
        count = self.__app.Workbooks.Count
        strName = strName.replace('/', '\\')
        for i in range(count):
            wb = self.__app.Workbooks.Item(i+1)
            if wb.FullName == strName:
                return xlbWorkBook(self.__app, wb)
        return None

    #-------------------------------------------------------------------------------
    def quit(self):
        self.__app.Quit()

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbWorkBook(object):
    #-------------------------------------------------------------------------------
    def __init__(self, app, wb):
        self.__app = app
        self.__wb  = wb

    #-------------------------------------------------------------------------------
    def save(self):
        self.__wb.Save()

    #-------------------------------------------------------------------------------
    def saveAs(self, strName):
        self.__wb.SaveAs(strName, self.__app.DefaultSaveFormat)

    #-------------------------------------------------------------------------------
    def getName(self):
        return self.__wb.FullName

    #-------------------------------------------------------------------------------
    def addWorkSheet(self):
        ws = self.__wb.Worksheets.Add()
        return xlbWorkSheet(self.__app, ws)

    #-------------------------------------------------------------------------------
    def getWorkSheet(self, strName):
        count = self.__wb.Worksheets.Count
        for i in range(count):
            ws = self.__wb.Worksheets.Item(i+1)
            if ws.Name == strName:
                return xlbWorkSheet(self, ws)
        return None

    #-------------------------------------------------------------------------------
    def close(self):
        self.__wb.Close()

    #-------------------------------------------------------------------------------
    def closeNoSave(self):
        self.__wb.Saved = True
        self.close()

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbWorkSheet(object):
    #-------------------------------------------------------------------------------
    def __init__(self, wb, ws):
        self.__wb = wb
        self.__ws = ws

    #-------------------------------------------------------------------------------
    def getName(self):
        return self.__ws.Name

    #-------------------------------------------------------------------------------
    def setName(self, strName):
        self.__ws.Name = strName

    #-------------------------------------------------------------------------------
    def getRange(self, iFrmRow, iFrmCol, iToRow=0, iToCol=0):
        range = self.__ws.Range(self.__getRangeId(iFrmRow, iFrmCol, iToRow, iToCol))
        return xlbRange(self, range)

    #-------------------------------------------------------------------------------
    def addChart(self, xlRange, chartType):
        chart = self.__ws.Shapes.AddChart().Chart
        chart.ChartType = chartType        
        chart.SetSourceData(xlRange.getRaw())
        chart.Name = "test"
        return xlbChart(self, chart)

    #-------------------------------------------------------------------------------
    def getChart(self, strName):
        chart = self.__ws.ChartObjects(strName)
        if chart == None:
            raise Exception('Could not find chart ( %s )' % strName)
        return xlbChart(self, chart)

    #-------------------------------------------------------------------------------
    def __getRangeId(self, iFrmRow, iFrmCol, iToRow=0, iToCol=0):
        if iFrmRow == 0 or iFrmCol == 0:
            return ''

        strFrmCell =  self.__getCellId(iFrmRow, iFrmCol)         
        if iToRow == 0 or iToCol == 0:
            return strFrmCell + ":" + strFrmCell

        strToCell =  self.__getCellId(iToRow, iToCol)        
        return strFrmCell + ":" + strToCell   

    #-------------------------------------------------------------------------------
    def __getCellId(self, iRow, iCol):
        strCol = ''
        iDividend = iCol
        iModulo   = 0
        while iDividend > 0:
            iModulo   = (iDividend - 1) % 26
            strCol    = chr(65 + iModulo) + strCol
            iDividend = (int)((iDividend - iModulo) / 26)
        strCell = strCol + str(iRow)
        return strCell

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbRange(object):
    #-------------------------------------------------------------------------------
    def __init__(self, ws, range):
        self.__ws = ws
        self.__range = range

    #-------------------------------------------------------------------------------
    def getRaw(self):
        return self.__range

    #-------------------------------------------------------------------------------
    def setVals(self, lstVals):
        self.__range.Value = lstVals
        self.__range.HorizontalAlignment = win32com.client.constants.xlLeft

    #-------------------------------------------------------------------------------
    def getVals(self):
        return self.__range.Value

    #---------------------------------------------------------------------------
    def setArray(self, arData):
        import numpy as np

        lstVals = []
        for iRow in range(arData.shape[0]):
            lstColVals = []
            tplRowData = arData[iRow]
            if len(arData.shape) == 1:
                tplRowData = tuple([tplRowData])
            for iCol in range(len(tplRowData)):
                val = tplRowData[iCol]
                
                if isinstance(val, types.StringType) or isinstance(val, types.UnicodeType):
                    cellVal = val
                elif val == None:
                    cellVal = ''
                elif math.isnan(val) or math.isinf(val):
                    cellVal =''
                else:
                    cellVal = np.asscalar(val)
                
                lstColVals.append(cellVal)                     
            lstVals.append(lstColVals)

        self.setVals(lstVals)

    #---------------------------------------------------------------------------
    def getArray(self):
        import numpy as np

        return np.array(self.getVals())

    #---------------------------------------------------------------------------
    def setFillColor(self, color):
        self.__range.Interior.Color = color

    #---------------------------------------------------------------------------
    def setFontColor(self, color):
        self.__range.Font.Color = color

    #---------------------------------------------------------------------------
    def setBorder(self, borderType):
        self.__range.Borders(win32com.client.constants.xlEdgeTop).Weight    = borderType
        self.__range.Borders(win32com.client.constants.xlEdgeBottom).Weight = borderType
        self.__range.Borders(win32com.client.constants.xlEdgeRight).Weight  = borderType
        self.__range.Borders(win32com.client.constants.xlEdgeLeft).Weight   = borderType

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbChart(object):
    #-------------------------------------------------------------------------------
    def __init__(self, ws, chart):
        self.__ws = ws
        self.__chart = chart

    #-------------------------------------------------------------------------------
    def setName(self, strName):
        self.__chart.Name = strName

# End of file
