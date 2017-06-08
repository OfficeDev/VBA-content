---
title: Cell.Selected Property (Publisher)
keywords: vbapb10.chm5111832
f1_keywords:
- vbapb10.chm5111832
ms.prod: publisher
api_name:
- Publisher.Cell.Selected
ms.assetid: b07f40bf-a14b-9b2a-2e0d-dc907cc78748
ms.date: 06/08/2017
---


# Cell.Selected Property (Publisher)

Returns  **True** if a cell is selected. Read-only **Boolean**.


## Syntax

 _expression_. **Selected**

 _expression_A variable that represents a  **Cell** object.


## Example

This example determines if a cell in the specified table is selected and if it is, enters text into the cell.


```vb
Sub IsCellSelected() 
 Dim cel As Cell 
 With ActiveDocument.Pages(1).Shapes(1).Table 
 For Each cel In .Cells 
 If cel.Selected Then 
 cel.TextRange.Text = "This cell is selected." 
 End If 
 Next cel 
 End With 
End Sub
```


