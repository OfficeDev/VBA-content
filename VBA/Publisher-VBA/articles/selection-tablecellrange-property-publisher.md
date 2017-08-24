---
title: Selection.TableCellRange Property (Publisher)
keywords: vbapb10.chm851975
f1_keywords:
- vbapb10.chm851975
ms.prod: publisher
api_name:
- Publisher.Selection.TableCellRange
ms.assetid: d683e830-6bcd-4b53-844b-605fab184a4c
ms.date: 06/08/2017
---


# Selection.TableCellRange Property (Publisher)

Returns a  **CellRange** object that represents the cells in a table selection.


## Syntax

 _expression_. **TableCellRange**

 _expression_A variable that represents a  **Selection** object.


### Return Value

CellRange


## Example

This example fills the table cells in a selection.


```vb
Sub FillTableCellRange() 
 Dim intCount As Integer 
 With Selection 
 If .Type = pbSelectionTableCells Then 
 With .TableCellRange 
 For intCount = 1 To .Count 
 .Item(intCount).Fill.ForeColor.RGB = RGB _ 
 (Red:=0, Green:=255, Blue:=255) 
 Next intCount 
 End With 
 End If 
 End With 
End Sub
```


