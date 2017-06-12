---
title: Rows Object (PowerPoint)
keywords: vbapp10.chm625000
f1_keywords:
- vbapp10.chm625000
ms.prod: powerpoint
api_name:
- PowerPoint.Rows
ms.assetid: 9a72b6bb-2aec-e37b-f1a2-005f910e1335
ms.date: 06/08/2017
---


# Rows Object (PowerPoint)

A collection of  **[Row](http://msdn.microsoft.com/library/df5ca5df-8119-1af8-b698-d96669ed0a02%28Office.15%29.aspx)** objects that represent the rows in a table.


## Example

Use the [Rows](http://msdn.microsoft.com/library/f7003d61-62d4-8d00-15c5-d9a2c5d57625%28Office.15%29.aspx)property to return the  **Rows** collection. This example changes the height of all rows in the specified table to 160 points.


```
Dim i As Integer

With ActivePresentation.Slides(2).Shapes(4).Table

    For i = 1 To .Rows.Count

        .Rows.Height = 160

    Next i

End With
```

Use the [Add](http://msdn.microsoft.com/library/7cc0c530-e817-1983-0946-90e499470668%28Office.15%29.aspx)method to add a row to a table. This example inserts a row before the second row in the referenced table.




```
ActivePresentation.Slides(2).Shapes(5).Table.Rows.Add (2)
```

Use  **Rows** (index), where index is a number that represents the position of the row in the table, to return a single **Row** object. This example deletes the first row from the table in shape five on slide two.




```
ActivePresentation.Slides(2).Shapes(5).Table.Rows(1).Delete
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/7cc0c530-e817-1983-0946-90e499470668%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/34a6d828-4c5e-098b-2c34-71b7cea0e9e2%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/e180bd7c-5ac2-72eb-4793-b08e0ea7eb3a%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/bfb443ea-abe0-401e-3aa9-ff47aa081c13%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/4bb27136-518a-3f51-6210-84caffd911d2%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
