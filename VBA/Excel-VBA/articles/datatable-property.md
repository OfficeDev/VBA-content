---
title: DataTable Property
keywords: vbagr10.chm66931
f1_keywords:
- vbagr10.chm66931
ms.prod: excel
api_name:
- Excel.DataTable
ms.assetid: bf432a3e-dd5e-db5b-63b3-4d037976edcc
ms.date: 06/08/2017
---


# DataTable Property

Returns a  **[DataTable](datatable-object.md)** object that represents the chart data table. Read-only.


## Example

This example adds a data table with an outline border to the chart.


```vb
With myChart 
 .HasDataTable = True 
 .DataTable.HasBorderOutline = True 
End With
```


