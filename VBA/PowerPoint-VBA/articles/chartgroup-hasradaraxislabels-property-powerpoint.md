---
title: ChartGroup.HasRadarAxisLabels Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.HasRadarAxisLabels
ms.assetid: ae8102a3-db43-410e-06fe-ab9f7f7ab6ff
ms.date: 06/08/2017
---


# ChartGroup.HasRadarAxisLabels Property (PowerPoint)

 **True** if a radar chart has axis labels. Read/write **Boolean**.


## Syntax

 _expression_. **HasRadarAxisLabels**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

This property applies only to radar charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables radar axis labels for chart group one of the first chart in the active document and sets their color. You should run the example on a radar chart.




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

