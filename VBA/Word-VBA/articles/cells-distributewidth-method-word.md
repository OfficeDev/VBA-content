---
title: Cells.DistributeWidth Method (Word)
keywords: vbawd10.chm155844815
f1_keywords:
- vbawd10.chm155844815
ms.prod: word
api_name:
- Word.Cells.DistributeWidth
ms.assetid: b617deaf-b84a-eed1-176d-9d38f2d10db8
ms.date: 06/08/2017
---


# Cells.DistributeWidth Method (Word)

Adjusts the width of the specified cells so that they are equal.


## Syntax

 _expression_ . **DistributeWidth**

 _expression_ Required. A variable that represents a **[Cells](cells-object-word.md)** collection.


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


[Cells Collection Object](cells-object-word.md)

