---
title: PivotTable.CubeFields Property (Excel)
keywords: vbaxl10.chm235132
f1_keywords:
- vbaxl10.chm235132
ms.prod: excel
api_name:
- Excel.PivotTable.CubeFields
ms.assetid: 043d6946-4d78-ba59-bef7-5aa4d000041d
ms.date: 06/08/2017
---


# PivotTable.CubeFields Property (Excel)

Returns the  **[CubeFields](cubefields-object-excel.md)** collection. Each **[CubeField](cubefield-object-excel.md)** object contains the properties of the cube field element. Read-only.


## Syntax

 _expression_ . **CubeFields**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example creates a list of cube field names for the data fields in the first OLAP-based PivotTable report on Sheet1.


```vb
Set objNewSheet = Worksheets.Add 
objNewSheet.Activate 
intRow = 1 
For Each objCubeFld In Worksheets("Sheet1").PivotTables(1).CubeFields 
 If objCubeFld.Orientation = xlDataField Then 
 objNewSheet.Cells(intRow, 1).Value = objCubeFld.Name 
 intRow = intRow + 1 
 End If 
Next objCubeFld
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

