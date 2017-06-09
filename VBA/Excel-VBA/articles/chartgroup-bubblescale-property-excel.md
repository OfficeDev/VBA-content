---
title: ChartGroup.BubbleScale Property (Excel)
keywords: vbaxl10.chm568095
f1_keywords:
- vbaxl10.chm568095
ms.prod: excel
api_name:
- Excel.ChartGroup.BubbleScale
ms.assetid: cbab742e-4e60-2d10-e8ec-0dcd2a5ff72a
ms.date: 06/08/2017
---


# ChartGroup.BubbleScale Property (Excel)

Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. Read/write  **Long** .


## Syntax

 _expression_ . **BubbleScale**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example sets the bubble size in chart group one to 200% of the default size.


```vb
With Worksheets(1).ChartObjects(1).Chart 
 .ChartGroups(1).BubbleScale = 200 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)

