---
title: Legend.IncludeInLayout Property (Word)
keywords: vbawd10.chm147196274
f1_keywords:
- vbawd10.chm147196274
ms.prod: word
api_name:
- Word.Legend.IncludeInLayout
ms.assetid: dd0e4c44-ba2a-191b-fa0a-d231a27506f9
ms.date: 06/08/2017
---


# Legend.IncludeInLayout Property (Word)

 **True** if a legend will occupy the chart layout space when a chart layout is being determined. The default is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeInLayout**

 _expression_ A variable that represents a **[Legend](legend-object-word.md)** object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title by using the  **Above Chart** command, the chart will resize smaller, as in previous versions of Microsoft Office. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


#### Concepts


[Legend Object](legend-object-word.md)

