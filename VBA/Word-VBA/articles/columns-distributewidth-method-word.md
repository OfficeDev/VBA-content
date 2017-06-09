---
title: Columns.DistributeWidth Method (Word)
keywords: vbawd10.chm155910347
f1_keywords:
- vbawd10.chm155910347
ms.prod: word
api_name:
- Word.Columns.DistributeWidth
ms.assetid: 91123d8e-faf0-79e5-ecc4-fabe68911b6c
ms.date: 06/08/2017
---


# Columns.DistributeWidth Method (Word)

Adjusts the width of the specified columns so that they are equal.


## Syntax

 _expression_ . **DistributeWidth**

 _expression_ Required. A variable that represents a **[Columns](columns-object-word.md)** collection.


## Example

This example adjusts the width of the columns in the first table in the active document so that they're equal.


```vb
ActiveDocument.Tables(1).Columns.DistributeWidth
```

This example adjusts the height of the selected cells.




```vb
If Selection.Cells.Count >= 2 Then 
 Selection.Cells.DistributeWidth 
End If
```


## See also


#### Concepts


[Columns Collection Object](columns-object-word.md)

