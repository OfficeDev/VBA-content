---
title: ChartGroup.UpBars Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.UpBars
ms.assetid: c6496698-c9a3-eee4-f829-f2feec787118
ms.date: 06/08/2017
---


# ChartGroup.UpBars Property (PowerPoint)

Returns the up bars on a line chart. Read-only  **[UpBars](upbars-object-powerpoint.md)**.


## Syntax

 _expression_. **UpBars**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

This property applies only to line charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables up and down bars for chart group one of the first chart in the active document, and then sets their colors. You should run the example on a 2-D line chart that contains two series that cross each other at one or more data points.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasUpDownBars = True

            .DownBars.Interior.ColorIndex = 3

            .UpBars.Interior.ColorIndex = 5

        End With

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

