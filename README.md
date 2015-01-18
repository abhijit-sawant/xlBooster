xlBooster
==========

Object oriented python interface to Excel. Users can manipulate Excel tabels from python. It also provides integration with 
NumPy so that one can push NumPy arrays to Excel table. This interface can be used even if NumPy is not installed. In that
case only NumPy related features will not be available.

Manipulate Excel table from python
----------------------------------

Set excel table values using python list

    import xlb
    
    app = xlb.xlbApp()
    wb = app.addWorkBook()
    ws = wb.addWorkSheet()
    ws.activate()
    
    #get range starting from cell 1,1 to 2,2
    r = ws.getRange(1, 1, 2, 2)
    
    #set values on range
    l = [[1, 11], [2, 22]]
    r.setList(l)
