---
title: ChartGroup.DoughnutHoleSize Property (PowerPoint)
keywords: vbapp10.chm692009
f1_keywords:
- vbapp10.chm692009
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.DoughnutHoleSize
ms.assetid: bd5fab99-265b-e9d9-3cb4-63d7e270d8b1
ms.date: 06/08/2017
---


# ChartGroup.DoughnutHoleSize Property (PowerPoint)

Returns or sets the size of the hole in a doughnut chart group. Read/write  **Long**.


## Syntax

 _expression_. **DoughnutHoleSize**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

The hole size is expressed as a percentage of the chart size, from 10 through 90 percent.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the hole size for doughnut group one of the first chart in the active document. You should run the example on a 2-D doughnut chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.DoughnutGroups(1).DoughnutHoleSize = 10

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

