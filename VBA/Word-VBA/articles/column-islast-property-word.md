---
title: Column.IsLast Property (Word)
keywords: vbawd10.chm156172293
f1_keywords:
- vbawd10.chm156172293
ms.prod: word
api_name:
- Word.Column.IsLast
ms.assetid: 9f5e51fe-4bb7-a179-4dde-373f7798f200
ms.date: 06/08/2017
---


# Column.IsLast Property (Word)

 **True** if the specified column or row is the last one in the table. Read-only **Boolean** .


## Syntax

 _expression_ . **IsLast**

 _expression_ Required. A variable that represents a **[Column](column-object-word.md)** object.


## Example

This example determines whether the first column in the selection is the last column in the table.


```vb
If Selection.Information(wdWithInTable) = True Then 
 MsgBox Selection.Columns(1).IsLast 
End If
```


## See also


#### Concepts


[Column Object](column-object-word.md)

