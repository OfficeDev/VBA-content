---
title: ChartGroup.HasDropLines Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.HasDropLines
ms.assetid: d957d6c6-acde-7ef0-9786-6f0f32d29253
ms.date: 06/08/2017
---


# ChartGroup.HasDropLines Property (PowerPoint)

 **True** if the line chart or area chart has drop lines. Read/write **Boolean**.


## Syntax

 _expression_. **HasDropLines**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

This property applies only to line and area charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables drop lines for chart group one of the first chart in the active document and then sets their line style, weight, and color. You should run the example on a 2-D line chart that has one series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasDropLines = True

            With .DropLines.Border

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

