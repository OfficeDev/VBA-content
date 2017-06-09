---
title: Gridlines Object
keywords: vbagr10.chm131203
f1_keywords:
- vbagr10.chm131203
ms.prod: excel
api_name:
- Excel.Gridlines
ms.assetid: 8879cdea-609f-5994-3fb6-3a9d5fa849b4
ms.date: 06/08/2017
---


# Gridlines Object

Represents major or minor gridlines on the specified chart axis. Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers. This object isn't a collection. There's no object that represents a single gridline; either you have all gridlines for an axis turned on or you have them all turned off.


## Using the Gridlines Object

Use the  **MajorGridlines** property to return the **GridLines** object that represents the major gridlines for the axis. Use the **MinorGridlines** property to return the **GridLines** object that represents the minor gridlines for the axis. It's possible to return both major and minor gridlines at the same time.

The following example turns on major gridlines for the category axis on the chart and then formats the gridlines to be blue dashed lines.




```vb
With myChart.Axes(xlCategory) 
 .HasMajorGridlines = True 
 .MajorGridlines.Border.Color = RGB(0, 0, 255) 
 .MajorGridlines.Border.LineStyle = xlDash 
End With
```


