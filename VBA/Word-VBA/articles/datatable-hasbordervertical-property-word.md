---
title: DataTable.HasBorderVertical Property (Word)
keywords: vbawd10.chm46399492
f1_keywords:
- vbawd10.chm46399492
ms.prod: word
api_name:
- Word.DataTable.HasBorderVertical
ms.assetid: cc104c8c-73b2-00a1-2ea9-641215560637
ms.date: 06/08/2017
---


# DataTable.HasBorderVertical Property (Word)

 **True** if the chart data table has vertical cell borders. Read/write **Boolean** .


## Syntax

 _expression_ . **HasBorderVertical**

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

