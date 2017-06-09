---
title: Cell.AutoSum Method (Word)
keywords: vbawd10.chm156106958
f1_keywords:
- vbawd10.chm156106958
ms.prod: word
api_name:
- Word.Cell.AutoSum
ms.assetid: 5f8c36c3-2e26-8e5f-16c4-49d4c04144c1
ms.date: 06/08/2017
---


# Cell.AutoSum Method (Word)

Inserts an = (Formula) field that calculates and displays the sum of the values in table cells above or to the left of the cell specified in the expression.


## Syntax

 _expression_ . **AutoSum**

 _expression_ Required. A variable that represents a **[Cell](cell-object-word.md)** object.


## Remarks

For information about how Word determines which values to add, see the  **[Formula](cell-formula-method-word.md)** method.


## Example

This example creates a 3x3 table in a new document and sums the numbers in the first column.


```vb
Dim docNew as Document 
Dim tableNew as Table 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Selection.Range, 3, 3) 
 
With tableNew 
 .Cell(1, 1).Range.InsertAfter "10" 
 .Cell(2, 1).Range.InsertAfter "15" 
 .Cell(3, 1).AutoSum 
End With
```

This example sums the numbers above or to the left of the cell that contains the insertion point.




```vb
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).AutoSum 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

