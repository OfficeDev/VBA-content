---
title: Chart.Walls Property (Excel)
keywords: vbaxl10.chm149151
f1_keywords:
- vbaxl10.chm149151
ms.prod: excel
api_name:
- Excel.Chart.Walls
ms.assetid: fbee1165-7602-4d77-e5b6-8a127783c96e
ms.date: 06/08/2017
---


# Chart.Walls Property (Excel)

Returns a  **[Walls](walls-object-excel.md)** object that represents the walls of the 3-D chart. Read-only.


## Syntax

 _expression_ . **Walls**

 _expression_ A variable that represents a **Chart** object.


## Example

This example sets the color of the wall border of Chart1 to red. The example should be run on a 3-D chart.


```vb
Charts("Chart1").Walls.Border.ColorIndex = 3
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

