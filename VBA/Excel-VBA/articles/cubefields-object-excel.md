---
title: CubeFields Object (Excel)
keywords: vbaxl10.chm669072
f1_keywords:
- vbaxl10.chm669072
ms.prod: excel
api_name:
- Excel.CubeFields
ms.assetid: cfb7b4f4-e9c3-45a3-daa4-fe4d3c52fb1f
ms.date: 06/08/2017
---


# CubeFields Object (Excel)

A collection of all  **[CubeField](cubefield-object-excel.md)** objects in a PivotTable report that is based on an OLAP cube. Each **CubeField** object represents a hierarchy or measure field from the cube.


## Example

Use the  **[CubeFields](pivottable-cubefields-property-excel.md)** property to return the **CubeFields** collection. The following example creates a list of cube field names of the data fields in the first OLAP-based PivotTable report on Sheet1.


```vb
Set objNewSheet = Worksheets.Add 
intRow = 1 
For Each objCubeFld In _ 
 Worksheets("Sheet1").PivotTables(1).CubeFields 
 If objCubeFld.Orientation = xlDataField Then 
 objNewSheet.Cells(intRow, 1).Value = objCubeFld.Name 
 intRow = intRow + 1 
 End If 
Next objCubeFld
```

Use  **CubeFields** ( _index_ ), where _index_ is the cube field's index number, to return a single **CubeField** object. The following example determines the name of the second cube field in the first PivotTable report on the active worksheet.




```
strAlphaName = _ 
 ActiveSheet.PivotTables(1).CubeFields(2).Name
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

