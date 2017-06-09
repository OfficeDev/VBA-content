---
title: Legend Object
keywords: vbagr10.chm131212
f1_keywords:
- vbagr10.chm131212
ms.prod: excel
api_name:
- Excel.Legend
ms.assetid: ed529b98-ad11-94b9-68d9-01e325cca58f
ms.date: 06/08/2017
---


# Legend Object

Represents the legend in the specified chart. Each chart can have only one legend. The  **Legend** object contains one or more **[LegendEntry](legendentry-object.md)** objects; each  **LegendEntry** object contains a **[LegendKey](legendkey-object.md)** object.


## Using the Legend Object

Use the  **Legend** property to return the **Legend** object. The following example sets the font style for the legend to bold.


```vb
myChart.Legend.Font.Bold = True
```


## Remarks

The chart legend isn't visible unless the  **[HasLegend](haslegend-property.md)** property is  **True**. If this property is  **False**, properties and methods of the  **Legend** object will fail.


