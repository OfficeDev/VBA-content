---
title: Table.SortAscending Method (Word)
keywords: vbawd10.chm156303372
f1_keywords:
- vbawd10.chm156303372
ms.prod: word
api_name:
- Word.Table.SortAscending
ms.assetid: 5a73ac7a-917d-7559-99c1-cb20f39b864d
ms.date: 06/08/2017
---


# Table.SortAscending Method (Word)

Sorts paragraphs or table rows in ascending alphanumeric order.


## Syntax

 _expression_ . **SortAscending**

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


## Remarks

The first table row is considered a header record and isn't included in the sort. Use the  **Sort** method to include the first row in a sort. This method offers a simplified form of sorting intended for mail merge data sources that contain columns of data. For most sorting tasks, use the **Sort** method.


## Example

This example sorts the table that contains the selection in ascending order.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Tables(1).SortAscending 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Table Object](table-object-word.md)

