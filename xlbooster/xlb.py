#!/usr/bin/env python

# Copyright 2015 Abhijit Sawant (abhijit.abhi1980@gmail.com)
#
# Licensed under the MIT License. You may not use this file except in compliance with the 
# License. You may obtain a copy of the License at
# 
# http://opensource.org/licenses/MIT


#try:
#    #from win32com.client import gencache
#    gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 6) #excel 2007
#    #gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 6) #excel 2013
#except:
#    raise Exception('Could not generate required Excel constatns. Import of module failed.')

import math
import types
import win32com.client
import xlbooster.constants as constants

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbApp(object):
    """Excel application class."""
    #-------------------------------------------------------------------------------
    def __init__(self):
        """The excel application will be launched on creation of this class's instance."""
        self.__app = win32com.client.DispatchEx("Excel.Application")
        self.__app.Visible = 1

    #-------------------------------------------------------------------------------
    def addWorkBook(self):
        """Adds new workbook and returns xlbWorkBook object."""
        wb = self.__app.Workbooks.Add()
        return xlbWorkBook(self.__app, wb)

    #-------------------------------------------------------------------------------
    def openWorkBook(self, strPath):
        """Opens workbook and returns xlbWorkBook object.
        
        strPath is a complete path including name.
        """
        wb = self.__app.Workbooks.Open(strPath)
        return xlbWorkBook(self.__app, wb)

    #-------------------------------------------------------------------------------
    def getWorkBook(self, strName):
        """Get xlbWorkBook object for already opened workbook

        strName is a complete path including name in case this workbook is already saved on disk.
        """
        count = self.__app.Workbooks.Count
        strName = strName.replace('/', '\\')
        for i in range(count):
            wb = self.__app.Workbooks.Item(i+1)
            if wb.FullName == strName:
                return xlbWorkBook(self.__app, wb)
        return None

    #-------------------------------------------------------------------------------
    def quit(self):
        """Quit the application."""
        self.__app.Quit()

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbWorkBook(object):
    """Workbook class."""
    #-------------------------------------------------------------------------------
    def __init__(self, app, wb):
        self.__app = app
        self.__wb  = wb

    #-------------------------------------------------------------------------------
    def save(self):
        """Save workbook with default name."""
        self.__wb.Save()

    #-------------------------------------------------------------------------------
    def saveAs(self, strName):
        """Save work book with provided name strName."""
        self.__wb.SaveAs(strName, self.__app.DefaultSaveFormat)

    #-------------------------------------------------------------------------------
    def getName(self):
        """Get full name of workbook."""
        return self.__wb.FullName

    #-------------------------------------------------------------------------------
    def addWorkSheet(self):
        """Add worksheet and return xlbWorkSheet object."""
        ws = self.__wb.Worksheets.Add()
        return xlbWorkSheet(self.__app, ws)

    #-------------------------------------------------------------------------------
    def getWorkSheet(self, strName):
        """Get worksheet with name strName."""
        count = self.__wb.Worksheets.Count
        for i in range(count):
            ws = self.__wb.Worksheets.Item(i+1)
            if ws.Name == strName:
                return xlbWorkSheet(self, ws)
        return None

    #-------------------------------------------------------------------------------
    def close(self):
        """Close workbook.
        
        If there are any unsaved changes Excel will pop open warning dialog.
        """
        self.__wb.Close()

    #-------------------------------------------------------------------------------
    def closeNoSave(self):
        """Close workbook without saving.
        
        Even if there are any unsaved changes Excel will not pop open warning dialog.
        These unsaved changes will be lost.
        """
        self.__wb.Saved = True
        self.close()

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbWorkSheet(object):
    """Worksheet class."""
    #-------------------------------------------------------------------------------
    def __init__(self, wb, ws):
        self.__wb = wb
        self.__ws = ws

    #-------------------------------------------------------------------------------
    def getName(self):
        """Get name of the worksheet."""
        return self.__ws.Name

    #-------------------------------------------------------------------------------
    def setName(self, strName):
        """Set name of the worksheet. """
        self.__ws.Name = strName

    #-------------------------------------------------------------------------------
    def getRange(self, iFrmRow, iFrmCol, iToRow=0, iToCol=0):
        """Create range object xlbRange.
        
        iFrmRow, iFromCol are row and column numbers for top left cell of range
        iToRow, iToCol are row and column numbers for bottom right cell of range
        """
        range = self.__ws.Range(self.__getRangeId(iFrmRow, iFrmCol, iToRow, iToCol))
        return xlbRange(self, range)

    #-------------------------------------------------------------------------------
    def addChart(self, xlRange, chartType):
        """Add chart to worksheet and return xlbChart object.

        xlRange is a range object for chart data.
        chartType is enum constant for chart type. The constants module provides enums which are exactly same as Excel VB.
        """
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
    """Range class."""
    #-------------------------------------------------------------------------------
    def __init__(self, ws, range):
        self.__ws = ws
        self.__range = range

    #-------------------------------------------------------------------------------
    def getRaw(self):
        """This method returns Excel range objbect.
        
        The object returned is a COM object. The direct use of this method by developer is NOT recommended.
        """
        return self.__range

    #-------------------------------------------------------------------------------
    def setVals(self, lstVals):
        """Set values of cells in range."""
        self.__range.Value = lstVals
        self.__range.HorizontalAlignment = constants.xlLeft

    #-------------------------------------------------------------------------------
    def getVals(self):
        """Get values of cells in range as a list."""
        return self.__range.Value

    #---------------------------------------------------------------------------
    def setArray(self, arData):
        """Set values of cells in range from NumPy array arData."""
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
        """Get values of cells in range as NumPy array."""
        import numpy as np

        return np.array(self.getVals())

    #---------------------------------------------------------------------------
    def setFillColor(self, color):
        """Set fill color of cells in range.
        
        color is a hex value.
        """
        self.__range.Interior.Color = color

    #---------------------------------------------------------------------------
    def setFontColor(self, color):
        """Set font color of cells in range.
        
        color is a hex value.
        """
        self.__range.Font.Color = color

    #---------------------------------------------------------------------------
    def setFont(self, name = '', style = ''):
        """Set font of cells in range."""
        self.__range.Font.Name      = name
        self.__range.Font.FontStyle = style

    #---------------------------------------------------------------------------
    def setBorder(self, borderType):
        """Set border type of range.
        
        borderType is a enum for border type. The constants module provides enums which are exactly same as Excel VB.
        """
        self.__range.Borders(constants.xlEdgeTop).Weight    = borderType
        self.__range.Borders(constants.xlEdgeBottom).Weight = borderType
        self.__range.Borders(constants.xlEdgeRight).Weight  = borderType
        self.__range.Borders(constants.xlEdgeLeft).Weight   = borderType

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
class xlbChart(object):
    """Chart class."""
    #-------------------------------------------------------------------------------
    def __init__(self, ws, chart):
        self.__ws = ws
        self.__chart = chart

    #-------------------------------------------------------------------------------
    def setName(self, strName):
        self.__chart.Name = strName

# End of file
