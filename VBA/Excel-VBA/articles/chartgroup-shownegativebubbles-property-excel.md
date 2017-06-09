---
title: ChartGroup.ShowNegativeBubbles Property (Excel)
keywords: vbaxl10.chm568096
f1_keywords:
- vbaxl10.chm568096
ms.prod: excel
api_name:
- Excel.ChartGroup.ShowNegativeBubbles
ms.assetid: 1f1288d5-71c5-f5da-583c-584db90c6c33
ms.date: 06/08/2017
---


# ChartGroup.ShowNegativeBubbles Property (Excel)

 **True** if negative bubbles are shown for the chart group. Valid only for bubble charts. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowNegativeBubbles**

 _expression_ A variable that represents a **ChartGroup** object.


## Example


```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .ChartGroups(1).ShowNegativeBubbles = True
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)

