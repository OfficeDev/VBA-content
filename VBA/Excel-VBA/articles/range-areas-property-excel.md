---
title: Range.Areas Property (Excel)
keywords: vbaxl10.chm144081
f1_keywords:
- vbaxl10.chm144081
ms.prod: excel
api_name:
- Excel.Range.Areas
ms.assetid: 31fc03b4-25b6-27ae-2350-b34c6c6ba255
ms.date: 06/08/2017
---


# Range.Areas Property (Excel)

Returns an  **[Areas](areas-object-excel.md)** collection that represents all the ranges in a multiple-area selection. Read-only.


## Syntax

 _expression_ . **Areas**

 _expression_ A variable that represents a **Range** object.


## Remarks

For a single selection, the  **Areas** property returns a collection that contains one object — the original **Range** object itself. For a multiple-area selection, the **Areas** property returns a collection that contains one object for each selected area.


## Example

This example displays a message if the user tries to carry out a command when more than one area is selected. This example must be run from a worksheet.


```vb
If Selection.Areas.Count > 1 Then 
 MsgBox "Cannot do this to a multi-area selection." 
End If
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

