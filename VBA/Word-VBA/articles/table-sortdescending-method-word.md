---
title: Table.SortDescending Method (Word)
keywords: vbawd10.chm156303373
f1_keywords:
- vbawd10.chm156303373
ms.prod: word
api_name:
- Word.Table.SortDescending
ms.assetid: a72b25e9-06c2-8f2f-1dff-796768d43fff
ms.date: 06/08/2017
---


# Table.SortDescending Method (Word)

Sorts table rows in descending alphanumeric order.


## Syntax

 _expression_ . **SortDescending**

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


## Remarks

The first table row is considered a header record and isn't included in the sort. Use the  **Sort** method to include the header record in a sort.

This method offers a simplified form of sorting intended for mail-merge data sources that contain columns of data. For most sorting tasks, use the  **Sort** method.


## Example

This example creates a 5x5 table in a new document, inserts text into each cell, and then sorts the table in descending alphanumeric order.


```vb
Set newDoc = Documents.Add 
Set myTable = _ 
 newDoc.Tables.Add(Range:=Selection.Range, NumRows:=5, _ 
 NumColumns:=5) 
For iRow = 1 To myTable.Rows.Count 
 For iCol = 1 To myTable.Columns.Count 
 Set MyRange = myTable.Rows(iRow).Cells(iCol).Range 
 MyRange.InsertAfter "Cell" &; Str$(iRow) &; "," &; Str$(iCol) 
 Next iCol 
Next iRow 
MsgBox "Click OK to sort in descending order." 
myTable.SortDescending
```

This example sorts the table that contains the insertion point in descending alphanumeric order.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Tables(1).SortDescending 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Table Object](table-object-word.md)

