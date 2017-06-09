---
title: Row.Previous Property (Word)
keywords: vbawd10.chm156237929
f1_keywords:
- vbawd10.chm156237929
ms.prod: word
api_name:
- Word.Row.Previous
ms.assetid: 2f58f33e-f3da-613a-dbeb-370d35ff865b
ms.date: 06/08/2017
---


# Row.Previous Property (Word)

Returns a  **Row** object that represents the table row that is previous to the specified row. Read-only.


## Syntax

 _expression_ . **Previous**

 _expression_ A variable that represents a **[Row](row-object-word.md)** object.


## Example

If the selection is in a table, this example selects the contents of the previous row.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows(1).Previous.Select 
End If
```


## See also


#### Concepts


[Row Object](row-object-word.md)

