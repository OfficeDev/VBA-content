---
title: HasLinks Property
keywords: vbagr10.chm66630
f1_keywords:
- vbagr10.chm66630
ms.prod: excel
api_name:
- Excel.HasLinks
ms.assetid: 71e0e494-a96a-53e5-5e38-92b3ce331076
ms.date: 06/08/2017
---


# HasLinks Property

 **True** if the specified chart has links to an external data source. Read-only **Boolean**.


## Example

This example clears cells A1:D4 on the datasheet if the chart has no links.


```vb
With myChart.Application 
 If .HasLinks = False Then 
 .DataSheet.Range("A1:D4").Clear 
 End If 
End With
```


