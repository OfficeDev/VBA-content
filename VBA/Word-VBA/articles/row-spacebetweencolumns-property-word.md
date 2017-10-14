---
title: Row.SpaceBetweenColumns Property (Word)
keywords: vbawd10.chm156237830
f1_keywords:
- vbawd10.chm156237830
ms.prod: word
api_name:
- Word.Row.SpaceBetweenColumns
ms.assetid: 22b81246-e158-ace7-dbca-9fc277584c6e
ms.date: 06/08/2017
---


# Row.SpaceBetweenColumns Property (Word)

Returns or sets the distance (in points) between text in adjacent columns of the specified row or rows. Read/write  **Single** .


## Syntax

 _expression_ . **SpaceBetweenColumns**

 _expression_ Required. A variable that represents a **[Row](row-object-word.md)** object.


## Example

This example creates a 3x3 table in a new document and then sets the distance between columns in the first row to 0.5 inches.


```vb
Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Selection.Range, 3, 3) 
myTable.Rows(1).SpaceBetweenColumns = InchesToPoints(0.5)
```


## See also


#### Concepts


[Row Object](row-object-word.md)

