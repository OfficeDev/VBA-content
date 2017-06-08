---
title: Row.IsFirst Property (Word)
keywords: vbawd10.chm156237835
f1_keywords:
- vbawd10.chm156237835
ms.prod: word
api_name:
- Word.Row.IsFirst
ms.assetid: 5efc4afa-cd5d-e9f2-b77e-b1375fa258d7
ms.date: 06/08/2017
---


# Row.IsFirst Property (Word)

 **True** if the specified row is the first one in the table. Read-only **Boolean** .


## Syntax

 _expression_ . **IsFirst**

 _expression_ Required. A variable that represents a **[Row](row-object-word.md)** object.


## Example

This example determines whether the first row in the selection is the first row in the table.


```vb
MsgBox Selection.Rows(1).IsFirst
```


## See also


#### Concepts


[Row Object](row-object-word.md)

