---
title: Row Object (Publisher)
keywords: vbapb10.chm4915199
f1_keywords:
- vbapb10.chm4915199
ms.prod: publisher
api_name:
- Publisher.Row
ms.assetid: 11f4688b-b94e-fa09-7c1b-4cbcca330936
ms.date: 06/08/2017
---


# Row Object (Publisher)

Represents a row in a table. The  **Row** object is a member of the **[Rows](rows-object-publisher.md)** collection. The **Rows** collection includes all the rows in a specified table.
 


## Example

Use  **Rows** (index), where index is the row number, to return a single **Row** object. The index number represents the position of the row in the **Rows** collection (counting from left to right). This example selects the first row in the first shape on the second of the active publication. This example assumes the specified shape is a table and not another type of shape.
 

 

```
Sub SelectRow() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(1).Cells.Select 
End Sub
```

Use the  **[Item](rows-item-method-publisher.md)** method of a **[Rows](rows-object-publisher.md)** collection to return a **Row** object. This example sets the fill for all even numbered rows and clears the fill for all odd numbered rows in the specified table. This example assumes the specified shape is a table and not another type of shape.
 

 



```
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Row Mod 2 = 0 Then 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 Next celTable 
 Next rowTable 
 
End Sub
```

Use the  **[Add](rows-add-method-publisher.md)** method to add a row to a table. This example adds a row to the specified table on the second page of the active publication, and then adjusts the width, merges the cells, and sets the fill color. This example assumes the first shape is a table and not another type of shape.
 

 



```
Sub NewRow() 
 Dim rowNew As Row 
 
 Set rowNew = ActiveDocument.Pages(2).Shapes(1).Table.Rows _ 
 .Add(BeforeRow:=3) 
 With rowNew 
 .Height = 2 
 .Cells.Merge 
 .Cells(1).Fill.ForeColor.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
End Sub
```

Use the  **[Delete](row-delete-method-publisher.md)** method to delete a row from a table. This example deletes the row added in the above example.
 

 



```
Sub DeleteRow() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(3).Delete 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](row-delete-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](row-application-property-publisher.md)|
|[Cells](row-cells-property-publisher.md)|
|[Height](row-height-property-publisher.md)|
|[Parent](row-parent-property-publisher.md)|

