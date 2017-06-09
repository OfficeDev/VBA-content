---
title: PivotField.AutoShowCount Property (Excel)
keywords: vbaxl10.chm240117
f1_keywords:
- vbaxl10.chm240117
ms.prod: excel
api_name:
- Excel.PivotField.AutoShowCount
ms.assetid: bbf7d754-04b3-d729-cf44-994fdc62db16
ms.date: 06/08/2017
---


# PivotField.AutoShowCount Property (Excel)

Returns the number of top or bottom items that are automatically shown in the specified PivotTable field. Read-only  **Long** .


## Syntax

 _expression_ . **AutoShowCount**

 _expression_ A variable that represents a **PivotField** object.


## Example

This example displays a message box showing the  **AutoShow** parameters for the Salesman field.


```vb
With Worksheets(1).PivotTables(1).PivotFields("salesman") 
 If .AutoShowType = xlAutomatic Then 
 r = .AutoShowRange 
 If r = xlTop Then 
 rn = "top" 
 Else 
 rn = "bottom" 
 End If 
 MsgBox "PivotTable report is showing " &; rn &; " " &; _ 
 .AutoShowCount &; " items in " &; .Name &; _ 
 " field by " &; .AutoShowField 
 Else 
 MsgBox "PivotTable report is not using AutoShow for this field" 
 End If 
End With
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

