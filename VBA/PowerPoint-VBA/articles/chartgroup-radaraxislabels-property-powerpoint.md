---
title: ChartGroup.RadarAxisLabels Property (PowerPoint)
keywords: vbapp10.chm692014
f1_keywords:
- vbapp10.chm692014
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.RadarAxisLabels
ms.assetid: 6bfef746-4616-7a63-0d1d-d0227a6e45f7
ms.date: 06/08/2017
---


# ChartGroup.RadarAxisLabels Property (PowerPoint)

Returns the radar axis labels for the specified chart group. Read-only  **[TickLabels](ticklabels-object-powerpoint.md)**.


## Syntax

 _expression_. **RadarAxisLabels**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables radar axis labels for chart group one for the first chart in the active document and then sets the color for the labels to red. You should run the example on a radar chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasRadarAxisLabels = True

            .RadarAxisLabels.Font.ColorIndex = 3

        End With

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

