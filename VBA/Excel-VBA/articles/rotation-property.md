---
title: Rotation Property
keywords: vbagr10.chm65595
f1_keywords:
- vbagr10.chm65595
ms.prod: excel
api_name:
- Excel.Rotation
ms.assetid: f78b6998-fae2-c80b-3a98-96ad359e6c47
ms.date: 06/08/2017
---


# Rotation Property

Returns or sets the rotation of the 3-D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3-D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3-D charts. Read/write  **Variant**.


## Example

This example sets the rotation of  `myChart` to 30 degrees. The example should be run on a 3-D chart.


```
myChart.Rotation = 30
```


