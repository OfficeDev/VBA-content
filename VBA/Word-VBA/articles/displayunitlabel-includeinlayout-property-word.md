---
title: DisplayUnitLabel.IncludeInLayout Property (Word)
keywords: vbawd10.chm94570866
f1_keywords:
- vbawd10.chm94570866
ms.prod: word
api_name:
- Word.DisplayUnitLabel.IncludeInLayout
ms.assetid: 05f119fe-d0b1-9309-f6d2-86abdd81c548
ms.date: 06/08/2017
---


# DisplayUnitLabel.IncludeInLayout Property (Word)

 **True** if a display unit label will occupy the chart layout space when a chart layout is being determined. The default is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeInLayout**

 _expression_ An expression that returns a **[DisplayUnitLabel](displayunitlabel-object-word.md)** object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title by using the  **Above Chart** command, the chart will resize smaller, as in previous versions of Microsoft Office. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


#### Concepts


[DisplayUnitLabel Object](displayunitlabel-object-word.md)

