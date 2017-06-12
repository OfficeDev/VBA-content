---
title: HasDataTable Property
keywords: vbagr10.chm5207470
f1_keywords:
- vbagr10.chm5207470
ms.prod: excel
api_name:
- Excel.HasDataTable
ms.assetid: 52d965ca-e4cf-35d5-0ac6-5a6144aedff0
ms.date: 06/08/2017
---


# HasDataTable Property

 **True** if the chart has a data table. Read/write **Boolean**.


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


