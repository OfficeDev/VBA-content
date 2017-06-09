---
title: CustomView.RowColSettings Property (Excel)
keywords: vbaxl10.chm508075
f1_keywords:
- vbaxl10.chm508075
ms.prod: excel
api_name:
- Excel.CustomView.RowColSettings
ms.assetid: 66e946bf-2f72-b7f4-a3fc-dd1ace044ec8
ms.date: 06/08/2017
---


# CustomView.RowColSettings Property (Excel)

 **True** if the custom view includes settings for hidden rows and columns (including filter information). Read-only **Boolean** .


## Syntax

 _expression_ . **RowColSettings**

 _expression_ A variable that represents a **CustomView** object.


## Example

This example creates a list of the custom views in the active workbook and their print settings and row and column settings.


```vb
With Worksheets(1) 
 .Cells(1,1).Value = "Name" 
 .Cells(1,2).Value = "Print Settings" 
 .Cells(1,3).Value = "RowColSettings" 
 rw = 0 
 For Each v In ActiveWorkbook.CustomViews 
 rw = rw + 1 
 .Cells(rw, 1).Value = v.Name 
 .Cells(rw, 2).Value = v.PrintSettings 
 .Cells(rw, 3).Value = v.RowColSettings 
 Next 
End With
```


## See also


#### Concepts


[CustomView Object](customview-object-excel.md)

