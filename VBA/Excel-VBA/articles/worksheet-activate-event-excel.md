---
title: Worksheet.Activate Event (Excel)
keywords: vbaxl10.chm502076
f1_keywords:
- vbaxl10.chm502076
ms.prod: excel
api_name:
- Excel.Worksheet.Activate
ms.assetid: 4fac262c-ea1a-1d2f-bd02-0537c843198c
ms.date: 06/08/2017
---


# Worksheet.Activate Event (Excel)

Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.


## Syntax

 _expression_ . **Activate**

 _expression_ A variable that represents a **Worksheet** object.


### Return Value

nothing


## Remarks

This event doesn't occur when you create a new window.

When you switch between two windows showing the same workbook, the WindowActivate event occurs, but the Activate event for the workbook doesn't occur.


## Example

This example sorts the range A1:A10 when the worksheet is activated.


```vb
Private Sub Worksheet_Activate() 
 Range("a1:a10").Sort Key1:=Range("a1"), Order:=xlAscending 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

