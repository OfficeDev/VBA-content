---
title: Selection.SplitTable Method (Word)
keywords: vbawd10.chm158663182
f1_keywords:
- vbawd10.chm158663182
ms.prod: word
api_name:
- Word.Selection.SplitTable
ms.assetid: 5d68a031-1927-ae5c-de11-963bca9c1d2c
ms.date: 06/08/2017
---


# Selection.SplitTable Method (Word)

Inserts an empty paragraph above the first row in the selection. .


## Syntax

 _expression_ . **SplitTable**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If the selection isn't in the first row of the table, the table is split into two tables. If the selection isn't in a table, an error occurs.


## Example

If the selection is in a table, this example splits the table.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.SplitTable 
End If
```

This example splits the first table in the active document between the first and second rows.




```vb
ActiveDocument.Tables(1).Rows(2).Select 
Selection.SplitTable
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

