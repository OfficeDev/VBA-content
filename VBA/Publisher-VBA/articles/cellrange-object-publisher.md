---
title: CellRange Object (Publisher)
keywords: vbapb10.chm5242879
f1_keywords:
- vbapb10.chm5242879
ms.prod: publisher
api_name:
- Publisher.CellRange
ms.assetid: 86e164f3-2a04-013f-3da8-d45c013eae7b
ms.date: 06/08/2017
---


# CellRange Object (Publisher)

A collection of  **[Cell](cell-object-publisher.md)** objects in a table column or row. The **CellRange** collection represents all the cells in the specified column or row.
 


## Remarks

Although the collection object is named  **CellRange** and is shown in the Object Browser, this keyword is not used in programming the Microsoft Publisher object model. The keyword **Cells** is used instead.
 

 
You cannot programmatically add to or delete individual cells from a Publisher table. Use the  **[AddTable](shapes-addtable-method-publisher.md)** method with the **[Shapes](shapes-object-publisher.md)** collection to add a new table to a publication. Use the **[Add](columns-add-method-publisher.md)** method of the **[Columns](columns-object-publisher.md)** or **[Rows](rows-object-publisher.md)** collections to add a column or row to a table. Use the **[Delete](column-delete-method-publisher.md)** method of the **Columns** or **Rows** collections to delete a column or row from a table.
 

 

## Example

Use the  **[Cells](column-cells-property-publisher.md)** property to return the **CellRange** collection. This example merges the cells in first column of the table.
 

 

```
Sub MergeCellsInFirstColumn() 
 With ActiveDocument.Pages(1).Shapes(1).Table 
 .Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=.Rows.Count, EndColumn:=1).Select 
 End With 
 Selection.TableCellRange.Merge 
End Sub
```

Use the  **[Count](cellrange-count-property-publisher.md)** property to return the number of cells in a row, column, table or selection. This example displays a message with the number of cells the specified table.
 

 



```
Sub NumberOfTableCells() 
 MsgBox ActiveDocument.Pages(1).Shapes(1).Table _ 
 .Cells.Count 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Item](cellrange-item-method-publisher.md)|
|[Merge](cellrange-merge-method-publisher.md)|
|[Select](cellrange-select-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](cellrange-application-property-publisher.md)|
|[Column](cellrange-column-property-publisher.md)|
|[Count](cellrange-count-property-publisher.md)|
|[Height](cellrange-height-property-publisher.md)|
|[Parent](cellrange-parent-property-publisher.md)|
|[Row](cellrange-row-property-publisher.md)|
|[Width](cellrange-width-property-publisher.md)|

