---
title: HasBorderVertical Property
keywords: vbagr10.chm5207458
f1_keywords:
- vbagr10.chm5207458
ms.prod: excel
api_name:
- Excel.HasBorderVertical
ms.assetid: ee6f449d-369c-1953-8540-b8baa4b281ab
ms.date: 06/08/2017
---


# HasBorderVertical Property

 **True** if the chart data table has vertical cell borders. Read/write **Boolean**.


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


