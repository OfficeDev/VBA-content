---
title: Column.Next Property (Word)
keywords: vbawd10.chm156172391
f1_keywords:
- vbawd10.chm156172391
ms.prod: word
api_name:
- Word.Column.Next
ms.assetid: fa2953dc-f5a6-ff58-9a64-42f865725ac7
ms.date: 06/08/2017
---


# Column.Next Property (Word)

Returns the next column in a collection of table columns. Read-only.


## Syntax

 _expression_ . **Next**

 _expression_ A variable that represents a **[Column](column-object-word.md)** object.


## Example

If the selection is in a table, this example selects the contents of the next table column.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Columns(1).Next.Select 
End If
```


## See also


#### Concepts


[Column Object](column-object-word.md)

