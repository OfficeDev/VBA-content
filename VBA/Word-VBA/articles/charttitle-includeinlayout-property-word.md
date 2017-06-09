---
title: ChartTitle.IncludeInLayout Property (Word)
keywords: vbawd10.chm65276274
f1_keywords:
- vbawd10.chm65276274
ms.prod: word
api_name:
- Word.ChartTitle.IncludeInLayout
ms.assetid: 5adcb002-5b23-cd5b-06ea-7680ed359653
ms.date: 06/08/2017
---


# ChartTitle.IncludeInLayout Property (Word)

 **True** if a chart title will occupy the chart layout space when a chart layout is being determined. The default is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeInLayout**

 _expression_ An expression that returns a **[ChartTitle](charttitle-object-word.md)** object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title by using the  **Above Chart** command, the chart will resize smaller, as in previous versions of Microsoft Office. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


#### Concepts


[ChartTitle Object](charttitle-object-word.md)

