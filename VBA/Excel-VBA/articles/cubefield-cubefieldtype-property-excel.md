---
title: CubeField.CubeFieldType Property (Excel)
keywords: vbaxl10.chm668073
f1_keywords:
- vbaxl10.chm668073
ms.prod: excel
api_name:
- Excel.CubeField.CubeFieldType
ms.assetid: 86847717-2906-6f92-36f4-668f932d2237
ms.date: 06/08/2017
---


# CubeField.CubeFieldType Property (Excel)

Indicates whether the OLAP cube field is a hierarchy field or a measure field. Can be one of the  **[XlCubeFieldType](xlcubefieldtype-enumeration-excel.md)** constants.


## Syntax

 _expression_ . **CubeFieldType**

 _expression_ A variable that represents a **CubeField** object.


## Example

This example creates a list of cube field names for the measure fields in the first OLAP-based PivotTable report on Sheet1.


```vb
Set objNewSheet = Worksheets.Add 
objNewSheet.Activate 
intRow = 1 
For Each objCubeFld in Worksheets("Sheet1").PivotTables(1).CubeFields 
 If objCubeFld.CubeFieldType = xlMeasure Then 
 objNewSheet.Cells(intRow, 1).Value = objCubeFld.Name 
 intRow = intRow + 1 
 End If 
Next objCubeFld
```


## See also


#### Concepts


[CubeField Object](cubefield-object-excel.md)

