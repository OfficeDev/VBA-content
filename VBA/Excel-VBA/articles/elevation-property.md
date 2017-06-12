---
title: Elevation Property
keywords: vbagr10.chm5207335
f1_keywords:
- vbagr10.chm5207335
ms.prod: excel
api_name:
- Excel.Elevation
ms.assetid: 5158f5d5-6900-f453-c4bc-7b52a1e42110
ms.date: 06/08/2017
---


# Elevation Property

Returns or sets the elevation of the 3-D chart view, in degrees. Read/write  **Long**.


## Remarks

The chart elevation is the height at which you view the chart, in degrees. The default is 15 for most chart types. The value of this property must be between -90 and 90, except for 3-D bar charts, where it must be between 0 and 44.


## Example

This example sets the chart elevation to 34 degrees. The example should be run on a 3-D chart (the  **Elevation** property fails on 2-D charts).


```
myChart.Elevation = 34
```


