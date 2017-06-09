---
title: Range.Count Property (Excel)
keywords: vbaxl10.chm144107
f1_keywords:
- vbaxl10.chm144107
ms.prod: excel
api_name:
- Excel.Range.Count
ms.assetid: 080cbbe7-056f-b21c-9004-171a6acce664
ms.date: 06/08/2017
---


# Range.Count Property (Excel)

Returns a  **Long** value that represents the number of objects in the collection.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents a **Range** object.


## Remarks

The  **Count** property is functionally the same as the **[CountLarge](range-countlarge-property-excel.md)** property, except that the **Count** property will generate an overflow error if the specified range has more than 2,147,483,647 cells (one less than 2048 columns). The **CountLarge** property, however, can handle ranges up to the maximum size for a worksheet, which is 17,179,869,184 cells.


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


[Range Object](range-object-excel.md)

