---
title: Areas.Count Property (Excel)
keywords: vbaxl10.chm197073
f1_keywords:
- vbaxl10.chm197073
ms.prod: excel
api_name:
- Excel.Areas.Count
ms.assetid: c3c91bed-d3dd-7ffd-94be-f61cc3b973b7
ms.date: 06/08/2017
---


# Areas.Count Property (Excel)

Returns a  **Long** value that represents the number of objects in the collection.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents an **Areas** object.


## Example

This example displays the number of columns in the selection on Sheet1. The code also tests for a multiple-area selection; if one exists, the code loops on the areas of the multiple-area selection.


```vb
Sub DisplayColumnCount() 
 Dim iAreaCount As Integer 
 Dim i As Integer 
 
 Worksheets("Sheet1").Activate 
 iAreaCount = Selection.Areas.Count 
 
 If iAreaCount <= 1 Then 
 MsgBox "The selection contains " &; Selection.Columns.Count &; " columns." 
 Else 
 For i = 1 To iAreaCount 
 MsgBox "Area " &; i &; " of the selection contains " &; _ 
 Selection.Areas(i).Columns.Count &; " columns." 
 Next i 
 End If 
End Sub
```


## See also


#### Concepts


[Areas Collection](areas-object-excel.md)

