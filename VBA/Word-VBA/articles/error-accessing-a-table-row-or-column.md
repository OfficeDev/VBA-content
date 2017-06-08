---
title: Error Accessing a Table Row or Column
ms.prod: word
ms.assetid: 0fa3199d-a1a4-fb29-49c9-59bcb1d8c98b
ms.date: 06/08/2017
---


# Error Accessing a Table Row or Column

When you try to access an individual row or column in a drawn table, a run-time error may occur if the table is not uniform. For example, the following instruction posts an error if the first table in the active document does not have the same number of rows in each column.


```vb
Sub RemoveTableBorders() 
 ActiveDocument.Tables(1).Rows(1).Borders.Enable = False 
End Sub
```


You can avoid this error by first selecting the cells in a column or row using the  **[SelectColumn](selection-selectcolumn-method-word.md)** method or the  **[SelectRow](selection-selectrow-method-word.md)** method. After the selection is made, use the  **[Cells](selection-cells-property-word.md)** property with the  **[Selection](selection-object-word.md)** object. The following example selects the first row in the first document table. The  **Cells** property is used to access the selected cells (all the cells in the first row) so that borders can be removed.




```vb
Sub RemoveTableBorders() 
 ActiveDocument.Tables(1).Cell(1, 1).Select 
 With Selection 
 .SelectRow 
 .Cells.Borders.Enable = False 
 End With 
End Sub
```

The following example selects the first column in the first document table. The  **For Each...Next** loop is used to add text to each cell in the selection (all the cells in the first column).



```vb
Sub AddTextToTableCells() 
 Dim intCell As Integer 
 Dim oCell As Cell 
 
 ActiveDocument.Tables(1).Cell(1, 1).Select 
 Selection.SelectColumn 
 intCell = 1 
 
 For Each oCell In Selection.Cells 
 oCell.Range.Text = "Cell " &; intCell 
 intCell = intCell + 1 
 Next oCell 
End Sub
```


