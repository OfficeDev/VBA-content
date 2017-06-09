---
title: PivotField.AutoShowRange Property (Excel)
keywords: vbaxl10.chm240116
f1_keywords:
- vbaxl10.chm240116
ms.prod: excel
api_name:
- Excel.PivotField.AutoShowRange
ms.assetid: b554867d-a78a-f26a-24b0-405f2d8a7c54
ms.date: 06/08/2017
---


# PivotField.AutoShowRange Property (Excel)

Returns  **xlTop** if the top items are shown automatically in the specified PivotTable field; returns **xlBottom** if the bottom items are shown. Read-only **Long** .


## Syntax

 _expression_ . **AutoShowRange**

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

