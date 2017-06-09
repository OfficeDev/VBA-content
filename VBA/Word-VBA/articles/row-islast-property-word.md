---
title: Row.IsLast Property (Word)
keywords: vbawd10.chm156237834
f1_keywords:
- vbawd10.chm156237834
ms.prod: word
api_name:
- Word.Row.IsLast
ms.assetid: f3520ca6-ddd1-eb5c-1243-27e47559d8e7
ms.date: 06/08/2017
---


# Row.IsLast Property (Word)

 **True** if the specified row is the last one in the table. Read-only **Boolean** .


## Syntax

 _expression_ . **IsLast**

 _expression_ Required. A variable that represents a **[Row](row-object-word.md)** object.


## Example

This example determines whether the second row is the last row in the table.


```vb
MsgBox ActiveDocument.Tables(1).Rows(2).IsLast
```


## See also


#### Concepts


[Row Object](row-object-word.md)

