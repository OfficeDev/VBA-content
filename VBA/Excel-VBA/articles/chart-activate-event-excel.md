---
title: Chart.Activate Event (Excel)
keywords: vbaxl10.chm500073
f1_keywords:
- vbaxl10.chm500073
ms.prod: excel
api_name:
- Excel.Chart.Activate
ms.assetid: 7b878d1b-3059-93cb-389a-a2633f613a4d
ms.date: 06/08/2017
---


# Chart.Activate Event (Excel)

Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.


## Syntax

 _expression_ . **Activate**

 _expression_ A variable that represents a **Chart** object.


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


[Chart Object](chart-object-excel.md)

