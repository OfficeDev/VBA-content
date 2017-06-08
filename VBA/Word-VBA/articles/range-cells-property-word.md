---
title: Range.Cells Property (Word)
keywords: vbawd10.chm157155385
f1_keywords:
- vbawd10.chm157155385
ms.prod: word
api_name:
- Word.Range.Cells
ms.assetid: aa081698-53d0-2234-5ec3-6e9a4091caef
ms.date: 06/08/2017
---


# Range.Cells Property (Word)

Returns a  **[Cells](cells-object-word.md)** collection that represents the table cells in a range. Read-only.


## Syntax

 _expression_ . **Cells**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example creates a 3x3 table and assigns a sequential cell number to each cell in the table.


```vb
Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Selection.Range, 3, 3) 
i = 1 
For Each c In myTable.Range.Cells 
 c.Range.InsertAfter "Cell " &; i 
 i = i + 1 
Next c
```


## See also


#### Concepts


[Range Object](range-object-word.md)

