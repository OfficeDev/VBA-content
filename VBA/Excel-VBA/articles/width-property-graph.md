---
title: Width Property (Graph)
keywords: vbagr10.chm3077602
f1_keywords:
- vbagr10.chm3077602
ms.prod: excel
ms.assetid: 715e889e-184e-5021-3ad9-029dd78e3147
ms.date: 06/08/2017
---


# Width Property (Graph)

As it applies to the Application object, the Width property determines the distance from the left edge of the application window to the right edge of the application window. For all other objects, the Width property, determines the width of the object. Read/write Double for all objects, except for the Chart object, which is read/write Variant.

 _expression_. **Width**

 _expression_ Required. An expression that returns one of the above objects.

 If the window is minimized, Application.Width is read-only and returns the width of the window icon.

## Example

This example sets the width of the chart.


```
myChart.Width = 360
```


