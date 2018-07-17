---
title: Workbook.Activate Event (Excel)
keywords: vbaxl10.chm503074
f1_keywords:
- vbaxl10.chm503074
ms.prod: excel
api_name:
- Excel.Workbook.Activate
ms.assetid: 74bb6d8c-aec8-7bb6-5c30-9a20f9a7afe8
ms.date: 06/08/2017
---


# Workbook.Activate Event (Excel)

Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.


## Syntax

 _expression_ . **Activate**

 _expression_ A variable that represents a **Workbook** object.


### Return Value

nothing


## Remarks

This event doesn't occur when you create a new window.

When you switch between two windows showing the same workbook, the WindowActivate event occurs, but the Activate event for the workbook doesn't occur.


## Example

This example sorts the range A1:A10 when the worksheet is activated.


```vb
Private Sub Worksheet_Activate() 
 Range("a1:a10").Sort Key1:=Range("a1"), Order1:=xlAscending 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

