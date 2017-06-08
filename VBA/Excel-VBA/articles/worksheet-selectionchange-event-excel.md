---
title: Worksheet.SelectionChange Event (Excel)
keywords: vbaxl10.chm502073
f1_keywords:
- vbaxl10.chm502073
ms.prod: excel
api_name:
- Excel.Worksheet.SelectionChange
ms.assetid: 183e2ca7-06b2-f689-1f77-182dbfbf1e1d
ms.date: 06/08/2017
---


# Worksheet.SelectionChange Event (Excel)

Occurs when the selection changes on a worksheet.


## Syntax

 _expression_ . **SelectionChange**( **_Target_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **[Range](range-object-excel.md)**|The new selected range.|

## Example

This example scrolls through the workbook window until the selection is in the upper-left corner of the window.


```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range) 
 With ActiveWindow 
 .ScrollRow = Target.Row 
 .ScrollColumn = Target.Column 
 End With 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

