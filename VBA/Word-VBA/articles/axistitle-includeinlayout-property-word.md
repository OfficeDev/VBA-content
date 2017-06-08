---
title: AxisTitle.IncludeInLayout Property (Word)
keywords: vbawd10.chm98240882
f1_keywords:
- vbawd10.chm98240882
ms.prod: word
api_name:
- Word.AxisTitle.IncludeInLayout
ms.assetid: be578a06-8a5f-80b5-79bd-ff2c0bee1311
ms.date: 06/08/2017
---


# AxisTitle.IncludeInLayout Property (Word)

 **True** if an axis title will occupy the chart layout space when a chart layout is being determined. The default is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeInLayout**

 _expression_ An expression that returns a **[AxisTitle](axistitle-object-word.md)** object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title by using the  **Above Chart** command, the chart will resize smaller, as in previous versions of Microsoft Office. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


#### Concepts


[AxisTitle Object](axistitle-object-word.md)

