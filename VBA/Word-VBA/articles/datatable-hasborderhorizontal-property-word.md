---
title: DataTable.HasBorderHorizontal Property (Word)
keywords: vbawd10.chm46399490
f1_keywords:
- vbawd10.chm46399490
ms.prod: word
api_name:
- Word.DataTable.HasBorderHorizontal
ms.assetid: d0e943dc-179b-c070-dd5b-2d58cc354b09
ms.date: 06/08/2017
---


# DataTable.HasBorderHorizontal Property (Word)

 **True** if the chart data table has horizontal cell borders. Read/write **Boolean** .


## Syntax

 _expression_ . **HasBorderHorizontal**

 _expression_ A variable that represents a **[DataTable](datatable-object-word.md)** object.


## Example

The following example causes the data table for the first chart in the active document to be displayed with an outline border and no cell borders.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .HasDataTable = True 
 With .DataTable 
 .HasBorderHorizontal = False 
 .HasBorderVertical = False 
 .HasBorderOutline = True 
 End With 
 End With 
 End If 
End With
```


## See also


#### Concepts


[DataTable Object](datatable-object-word.md)

