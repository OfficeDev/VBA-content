---
title: HasBorderOutline Property
keywords: vbagr10.chm5207454
f1_keywords:
- vbagr10.chm5207454
ms.prod: excel
api_name:
- Excel.HasBorderOutline
ms.assetid: b98fd5e2-fe84-1736-eb94-9e6e51ac49a6
ms.date: 06/08/2017
---


# HasBorderOutline Property

 **True** if the chart data table has outline borders. Read/write **Boolean**.


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


