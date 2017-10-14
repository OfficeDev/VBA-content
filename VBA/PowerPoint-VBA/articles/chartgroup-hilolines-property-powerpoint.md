---
title: ChartGroup.HiLoLines Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.HiLoLines
ms.assetid: 3b575e71-79c9-83d8-4c2d-dfc36480099f
ms.date: 06/08/2017
---


# ChartGroup.HiLoLines Property (PowerPoint)

Returns the high-low lines for a series on a line chart. Read-only  **[HiLoLines](hilolines-object-powerpoint.md)**.


## Syntax

 _expression_. **HiLoLines**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

This property applies only to line charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables high-low lines for chart group one of the first chart in the active document and then sets their line style, weight, and color. You should run the example on a 2-D line chart that has three series of stock-quote-like data (high-low-close).




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasHiLoLines = True

            With .HiLoLines.Border

                .LineStyle = xlThin

                .Weight = xlMedium

                .ColorIndex = 3

            End With

        End With

    End If

End With


```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

