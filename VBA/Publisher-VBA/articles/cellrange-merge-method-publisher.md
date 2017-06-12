---
title: CellRange.Merge Method (Publisher)
keywords: vbapb10.chm5177352
f1_keywords:
- vbapb10.chm5177352
ms.prod: publisher
api_name:
- Publisher.CellRange.Merge
ms.assetid: f097659c-d1b8-f2bb-c4fc-5efc2b7417dd
ms.date: 06/08/2017
---


# CellRange.Merge Method (Publisher)

Merges the specified table cells with one another. The result is a single table cell.


## Syntax

 _expression_. **Merge**

 _expression_A variable that represents a  **CellRange** object.


## Example

This example merges the first two cells in the first two rows of the specified table.


```vb
Sub MergeCells() 
 ActiveDocument.Pages(1).Shapes(2).Table _ 
 .Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=2, EndColumn:=2).Merge 
End Sub
```


