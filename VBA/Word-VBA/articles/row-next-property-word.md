---
title: Row.Next Property (Word)
keywords: vbawd10.chm156237928
f1_keywords:
- vbawd10.chm156237928
ms.prod: word
api_name:
- Word.Row.Next
ms.assetid: d74be2bd-5b12-8478-1a09-744571b0bd66
ms.date: 06/08/2017
---


# Row.Next Property (Word)

Returns a  **Row** object that represents the table row that is next in the collection of rows in a table. Read-only.


## Syntax

 _expression_ . **Next**

 _expression_ A variable that represents a **[Row](row-object-word.md)** object.


## Example

If the selection is in a table, this example selects the contents of the next table row.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Row(1).Next.Select 
End If
```


## See also


#### Concepts


[Row Object](row-object-word.md)

