---
title: PivotTable.NullString Property (Excel)
keywords: vbaxl10.chm235114
f1_keywords:
- vbaxl10.chm235114
ms.prod: excel
api_name:
- Excel.PivotTable.NullString
ms.assetid: f9d678d1-5e9f-8d3b-1f9a-73e8679ae499
ms.date: 06/08/2017
---


# PivotTable.NullString Property (Excel)

Returns or sets the string displayed in cells that contain null values when the  **[DisplayNullString](pivottable-displaynullstring-property-excel.md)** property is **True** . The default value is an empty string (""). Read/write **String** .


## Syntax

 _expression_ . **NullString**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example causes the PivotTable report to display "NA" in cells that contain null values.


```vb
With Worksheets(1).PivotTables("Pivot1") 
 .NullString = "NA" 
 .DisplayNullString = True 
End With
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

