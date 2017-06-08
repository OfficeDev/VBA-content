---
title: Table.Uniform Property (Word)
keywords: vbawd10.chm156303465
f1_keywords:
- vbawd10.chm156303465
ms.prod: word
api_name:
- Word.Table.Uniform
ms.assetid: a156bedf-5426-be4c-b961-84a038f9bfd6
ms.date: 06/08/2017
---


# Table.Uniform Property (Word)

 **True** if all the rows in a table have the same number of columns. Read-only **Boolean** .


## Syntax

 _expression_ . **Uniform**

 _expression_ An expression that returns a **[Table](table-object-word.md)** object.


## Example

This example creates a table that contains a split cell and then displays a message box that confirms that the table doesn't have the same number of columns for each row.


```vb
Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Selection.Range, 5, 5) 
myTable.Cell(3, 3).Split 1, 2 
If myTable.Uniform = False Then MsgBox "Table is not uniform"
```

This example determines whether the table that contains the selection has the same number of columns for each row.




```vb
If Selection.Information(wdWithInTable) = True Then 
 MsgBox Selection.Tables(1).Uniform 
End If
```


## See also


#### Concepts


[Table Object](table-object-word.md)

