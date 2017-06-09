---
title: DataTable Object (Word)
keywords: vbawd10.chm708
f1_keywords:
- vbawd10.chm708
ms.prod: word
api_name:
- Word.DataTable
ms.assetid: 4e6094ea-3d83-6ec0-9788-9d22b884beb2
ms.date: 06/08/2017
---


# DataTable Object (Word)

Represents a chart data table.


## Example

Use the  **[DataTable](chart-datatable-property-word.md)** property to return a **DataTable** object. The following example adds a data table with an outline border to embedded chart one.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.HasDataTable = True 
 .Chart.DataTable.HasBorderOutline = True 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


