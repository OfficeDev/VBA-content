---
title: Cell Object (Publisher)
keywords: vbapb10.chm5177343
f1_keywords:
- vbapb10.chm5177343
ms.prod: publisher
api_name:
- Publisher.Cell
ms.assetid: 5baafaa6-368e-9eae-30b9-90d2d89d5a5b
ms.date: 06/08/2017
---


# Cell Object (Publisher)

Represents a single table cell. The  **Cell** object is a member of the **[CellRange](http://msdn.microsoft.com/library/86e164f3-2a04-013f-3da8-d45c013eae7b%28Office.15%29.aspx)** collection. The **CellRange** collection represents all the cells in the specified object.


## Example

Use  **Cells** (index), where index is the cell number, to return a **Cell** object. This example merges the first two cells of the first column of the specified table.


```
Sub MergeCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(1) 
 .Cells(1).Merge MergeTo:=.Cells(2) 
 End With 
End Sub
```

This example applies a thick border around the first cell in the second column of the specified table.




```
Sub OutlineBorderCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(2).Cells(1) 
 .BorderLeft.Weight = 5 
 .BorderRight.Weight = 5 
 .BorderTop.Weight = 5 
 .BorderBottom.Weight = 5 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Merge](http://msdn.microsoft.com/library/09ae6910-ba47-4b0e-5792-ac9eb1298d57%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/37a026a3-67ec-7a13-5eb4-66e14918579d%28Office.15%29.aspx)|
|[Split](http://msdn.microsoft.com/library/99bc34df-c8dc-90e5-4262-dbe0a9c9b61d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/8ed632c6-ebcd-a6c6-3de0-42b40c17ffb4%28Office.15%29.aspx)|
|[BorderBottom](http://msdn.microsoft.com/library/78892893-a1c9-5151-fb7b-1449c01e0bd4%28Office.15%29.aspx)|
|[BorderDiagonal](http://msdn.microsoft.com/library/2c857a1b-2a0f-5796-9397-ad113dd984cb%28Office.15%29.aspx)|
|[BorderLeft](http://msdn.microsoft.com/library/f996a96f-4392-48c2-e5c2-bfe373a7997a%28Office.15%29.aspx)|
|[BorderRight](http://msdn.microsoft.com/library/da741816-d61c-61db-cf33-5b181780b902%28Office.15%29.aspx)|
|[BorderTop](http://msdn.microsoft.com/library/4119fcb7-7662-7ab5-ee56-4ef75aaa2766%28Office.15%29.aspx)|
|[CellTextOrientation](http://msdn.microsoft.com/library/ad2c2f15-358c-7bbc-b249-b886a99ea4a5%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/09e067a2-ee84-7a76-72b6-3b348238d020%28Office.15%29.aspx)|
|[Diagonal](http://msdn.microsoft.com/library/4ec93690-38ef-7434-55a5-419f14c9ea73%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/3ff3fda8-aca7-534e-6a56-99d6a3d9e9e2%28Office.15%29.aspx)|
|[HasText](http://msdn.microsoft.com/library/b44c5d24-7ac1-a63d-6986-05ed9c91dd8e%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/ced71ac0-eca8-0939-8812-fe0e79a47cba%28Office.15%29.aspx)|
|[MarginBottom](http://msdn.microsoft.com/library/a05fd3a4-f4d5-232a-1f5d-0fa1bce136bd%28Office.15%29.aspx)|
|[MarginLeft](http://msdn.microsoft.com/library/1b665a3b-6958-0548-ece1-9d3a7045eaac%28Office.15%29.aspx)|
|[MarginRight](http://msdn.microsoft.com/library/d297222e-7fc1-9225-e098-1a85d7734d77%28Office.15%29.aspx)|
|[MarginTop](http://msdn.microsoft.com/library/f408edd3-7199-b49a-817b-7b0e8461715c%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/2eecfc29-e349-4dfe-0751-b2c43dce2f7e%28Office.15%29.aspx)|
|[Row](http://msdn.microsoft.com/library/b961af2b-6b03-f54b-922e-d2e7633a3dfe%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/b07f40bf-a14b-9b2a-2e0d-dc907cc78748%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/31aa92d1-852f-3742-defa-94485411bcc3%28Office.15%29.aspx)|
|[VerticalTextAlignment](http://msdn.microsoft.com/library/793bf932-15d0-cce9-1d5d-aee5d260e1a0%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/894ab5ba-97a5-a731-cac2-151de813e5b8%28Office.15%29.aspx)|

