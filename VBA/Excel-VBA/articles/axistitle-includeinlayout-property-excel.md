---
title: AxisTitle.IncludeInLayout Property (Excel)
keywords: vbaxl10.chm566075
f1_keywords:
- vbaxl10.chm566075
ms.prod: excel
api_name:
- Excel.AxisTitle.IncludeInLayout
ms.assetid: ef84d235-6d60-f5c9-f185-e474a8b6a0e7
ms.date: 06/08/2017
---


# AxisTitle.IncludeInLayout Property (Excel)

 **True** if an axis title will occupy the chart layout space when a chart layout is being determined. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeInLayout**

 _expression_ A variable that represents an **AxisTitle** object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title using the  **Above Chart** command, the chart will resize smaller, as in Microsoft Office Excel 2003. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


#### Concepts


[AxisTitle Object](axistitle-object-excel.md)

