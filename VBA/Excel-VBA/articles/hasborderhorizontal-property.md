---
title: HasBorderHorizontal Property
keywords: vbagr10.chm67207
f1_keywords:
- vbagr10.chm67207
ms.prod: excel
api_name:
- Excel.HasBorderHorizontal
ms.assetid: 9d5a86ea-73f1-a149-8fc9-ce104cdb41a3
ms.date: 06/08/2017
---


# HasBorderHorizontal Property

 **True** if the chart data table has horizontal cell borders. Read/write **Boolean**.


## Example

This example causes the chart data table to be displayed with an outline border and no cell borders.


```vb
With myChart 
 .HasDataTable = True 
 With .DataTable 
 .HasBorderHorizontal = False 
 .HasBorderVertical = False 
 .HasBorderOutline = True 
 End With 
End With
```


