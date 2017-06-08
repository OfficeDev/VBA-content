---
title: Table Object (Publisher)
keywords: vbapb10.chm4849663
f1_keywords:
- vbapb10.chm4849663
ms.prod: publisher
api_name:
- Publisher.Table
ms.assetid: 09da4a0a-2230-067e-1cac-55321ea044c5
ms.date: 06/08/2017
---


# Table Object (Publisher)

Represents a single table.


## Example

Use the  **[Table](http://msdn.microsoft.com/library/a9b29d1f-2459-556c-56f8-f8f809b879c9%28Office.15%29.aspx)** property to return a **Table** object. The following example selects the specified table in the active publication.


```
Sub SelectTable() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = pbTable Then _ 
 .Table.Cells.Select 
 End With 
End Sub
```

Use the  **[AddTable](http://msdn.microsoft.com/library/1aa00f40-de41-12ed-8d4f-5e9c91cbf5af%28Office.15%29.aspx)** method to add a **Shape** object representing a table at the specified range. The following example adds a 5x5 table on the first page of the active publication, and then selects the first column of the new table.




```
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=5, NumColumns:=5, _ 
 Left:=72, Top:=300, Width:=400, Height:=100) 
 .Table.Columns(1).Cells.Select 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ApplyAutoFormat](http://msdn.microsoft.com/library/f792a5f3-0d1c-06de-a030-7a588ca372d2%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/9d808ec1-3f29-c2d4-b685-7acd3c6d0f18%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/42622697-aef1-0765-7d85-4919c298d92f%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/fb55ba62-64a4-2221-3cc7-b349dc2f6934%28Office.15%29.aspx)|
|[GrowToFitText](http://msdn.microsoft.com/library/d8822df7-a252-a5bb-be26-83df8ec5eb94%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e7c02be8-1888-4817-05bf-75b030e597fc%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/97a543b9-a1d7-c7f8-9f3c-e08256e0b364%28Office.15%29.aspx)|
|[TableDirection](http://msdn.microsoft.com/library/ffd664a8-781f-8fdc-055c-1ea7309b3b38%28Office.15%29.aspx)|

