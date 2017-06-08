---
title: BubbleScale Property
keywords: vbagr10.chm3076966
f1_keywords:
- vbagr10.chm3076966
ms.prod: excel
api_name:
- Excel.BubbleScale
ms.assetid: e3947690-3428-3f50-173b-b7889f9aac7f
ms.date: 06/08/2017
---


# BubbleScale Property

Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. Read/write Long.

 _expression_. **BubbleScale**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example sets the bubble size in chart group one to 200 percent of the default size.


```vb
With myChart 
 .ChartGroups(1).BubbleScale = 200 
End With
```


