xlBooster
==========

Object oriented python interface to Excel. Users can work with Excel tabel from python. It also provides integration with 
NumPy so that one can push NumPy array to Excel table or extract Excel tabel as NumPy array. This interface can be used even if NumPy is not installed. In that case only NumPy related features will not be available.

Interact with excel table using python list/tuple 

```python
import xlbooster.xlb as xlb

app = xlb.xlbApp()
wb = app.addWorkBook()
ws = wb.addWorkSheet()
ws.setName('Data')

#get range starting from cell 1,1 to 2,2
r = ws.getRange(1, 1, 2, 2)

#set values on range
l = [[1, 11], [2, 22]]
r.setList(l)

wb.saveAs('C:\\report')
wb.close()

#get values as tuple
wb = app.openWorkBook('C:\\report.xlsx')
ws = wb.getWorkSheet('Data')

r = ws.getRange(1, 1, 2, 2)
tup = r.getVals()
```

Interact with excel table using NumPy array

```python 
import numpy as np
import xlbooster.xlb as xlb

wb = app.addWorkBook()
ws = wb.getWorkSheet('Sheet1')

r = ws.getRange(1, 1, 2, 2)

#set values on range
ar_in = np.array([[1, 11], [2, 22]])
r.setArray(ar_in)

#get values as array
ar_out = r.getArray()
```
