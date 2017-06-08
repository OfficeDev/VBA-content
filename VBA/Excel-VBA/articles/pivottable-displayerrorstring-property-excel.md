---
title: PivotTable.DisplayErrorString Property (Excel)
keywords: vbaxl10.chm235104
f1_keywords:
- vbaxl10.chm235104
ms.prod: excel
api_name:
- Excel.PivotTable.DisplayErrorString
ms.assetid: 57ec3e1f-b6ea-dfd0-996e-6efa48bd9793
ms.date: 06/08/2017
---


# PivotTable.DisplayErrorString Property (Excel)

 **True** if the PivotTable report displays a custom error string in cells that contain errors. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayErrorString**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

Use the  **[ErrorString](pivottable-errorstring-property-excel.md)** property to set the custom error string.

This property is particularly useful for suppressing divide-by-zero errors when calculated fields are pivoted.


## Example

This example causes the PivotTable report to display a hyphen in cells that contain errors.


```vb
With Worksheets(1).PivotTables("Pivot1") 
 .ErrorString = "-" 
 .DisplayErrorString = True 
End With
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

