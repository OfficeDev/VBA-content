---
title: Column Object (Publisher)
keywords: vbapb10.chm5046271
f1_keywords:
- vbapb10.chm5046271
ms.prod: publisher
api_name:
- Publisher.Column
ms.assetid: 7f14fd4f-3919-8dd9-ed1e-988269b4b0c9
ms.date: 06/08/2017
---


# Column Object (Publisher)

Represents a single table column. The  **Column** object is a member of the **[Columns](columns-object-publisher.md)** collection. The **Columns** collection includes all the columns in a table, selection, or range.
 


## Example

Use  **Columns** (index), where index is the column number, to return a single **Column** object. The index number represents the position of the column in the **Columns** collection (counting from left to right). This example selects column three in the first shape in the active publication. This example assumes the first shape is a table and not another type of shape.
 

 

```
Sub SelectColumn() 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns(3).Cells.Select 
End Sub
```

Use the  **[Item](columns-item-method-publisher.md)** method of a **[Columns](columns-object-publisher.md)** collection to return a **Column** object. This example enters text into the first cell of the third column of the specified table and formats the text with a bold, 15-point font. This example assumes the first shape is a table and not another type of shape.
 

 



```
Sub ColumnHeading() 
 With ActiveDocument.Pages(2).Shapes(1).Table.Columns(3) _ 
 .Cells(1).Text 
 .Text = "Sales" 
 .Font.Bold = msoTrue 
 .Font.Size = 15 
 End With 
End Sub
```

Use the  **[Delete](column-delete-method-publisher.md)** method to delete a column from a table. This example deletes the column added in the above example.
 

 



```
Sub DeleteColumn() 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns(3).Delete 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](column-delete-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](column-application-property-publisher.md)|
|[Cells](column-cells-property-publisher.md)|
|[Parent](column-parent-property-publisher.md)|
|[Width](column-width-property-publisher.md)|

