---
title: PivotField.AutoShowField Property (Excel)
keywords: vbaxl10.chm240118
f1_keywords:
- vbaxl10.chm240118
ms.prod: excel
api_name:
- Excel.PivotField.AutoShowField
ms.assetid: 88d3a338-c809-0843-7968-9a8e60612445
ms.date: 06/08/2017
---


# PivotField.AutoShowField Property (Excel)

Returns the name of the data field used to determine the top or bottom items that are automatically shown in the specified PivotTable field. Read-only  **String** .


## Syntax

 _expression_ . **AutoShowField**

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

