---
title: PivotField.AutoShowType Property (Excel)
keywords: vbaxl10.chm240115
f1_keywords:
- vbaxl10.chm240115
ms.prod: excel
api_name:
- Excel.PivotField.AutoShowType
ms.assetid: a8146e5c-b1b4-7ff4-d2d7-bc98b863681d
ms.date: 06/08/2017
---


# PivotField.AutoShowType Property (Excel)

Returns  **xlAutomatic** if **AutoShow** is enabled for the specified PivotTable field; returns **xlManual** if **AutoShow** is disabled. Read-only **Long** .


## Syntax

 _expression_ . **AutoShowType**

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

