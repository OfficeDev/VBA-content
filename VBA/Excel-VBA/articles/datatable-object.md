---
title: DataTable Object
keywords: vbagr10.chm5207296
f1_keywords:
- vbagr10.chm5207296
ms.prod: excel
api_name:
- Excel.DataTable
ms.assetid: cf9aa637-3b5d-1e18-1956-291a0295dddf
ms.date: 06/08/2017
---


# DataTable Object

Represents a data table in the specified chart.


## Using the DataTable Object

Use the  **DataTable** property to return a **DataTable** object. The following example adds a data table with an outline border to the embedded chart.


```vb
With myChart 
 .HasDataTable = True 
 .DataTable.HasBorderOutline = True 
End With
```


