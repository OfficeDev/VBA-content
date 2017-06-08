---
title: LegendKey Object
keywords: vbagr10.chm5207614
f1_keywords:
- vbagr10.chm5207614
ms.prod: excel
api_name:
- Excel.LegendKey
ms.assetid: ab90cb64-1f81-dfcb-7542-cba68964acba
ms.date: 06/08/2017
---


# LegendKey Object

Represents a legend key in the specified chart legend. Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart. The legend key is linked to its associated series or trendline in such a way that changing the formatting of one simultaneously changes the formatting of the other.


## Using the LegendKey Object

Use the  **LegendKey** property to return the **LegendKey** object. The following example changes the marker background color to blue for the legend entry at the top of the legend in the chart. This simultaneously changes the formatting of every point in the series associated with this legend entry (if, that is, the associated series supports data markers).


```
myChart.Legend.LegendEntries(1) _ 
 .LegendKey.MarkerBackgroundColorIndex = 5
```


