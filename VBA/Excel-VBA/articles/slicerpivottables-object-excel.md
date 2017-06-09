---
title: SlicerPivotTables Object (Excel)
keywords: vbaxl10.chm910072
f1_keywords:
- vbaxl10.chm910072
ms.prod: excel
api_name:
- Excel.SlicerPivotTables
ms.assetid: 8302dc8a-3845-12b0-f88e-761f104f1dcc
ms.date: 06/08/2017
---


# SlicerPivotTables Object (Excel)

Represents information about the collection of PivotTables associated with the specified  **[SlicerCache](slicercache-object-excel.md)** object.


## Remarks

The  **SlicerPivotTables** collection contains information about the PivotTables the slicer cache is currently filtering. It provides properties for determining the number of PivotTables the slicer is associated with, and for retrieving **[PivotTable](pivottable-object-excel.md)** objects that represent the PivotTables being filtered. It also provides methods for adding and removing PivotTables from the **SlicerPivotTables** collection. The **SlicerPivotTables** collection will be empty if the slicer associated with the specified **SlicerCache** is not connected to any PivotTables.

 Use the **[PivotTables](slicercache-pivottables-property-excel.md)** property of the **SlicerCache** object to return the **SlicerPivotTables** collection associated with a **SlicerCache** , which in turn may be associated with one or more slicers.


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


