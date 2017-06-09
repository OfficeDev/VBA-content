---
title: ChartGroup.HasHiLoLines Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.HasHiLoLines
ms.assetid: 02122126-1ea9-0d94-ce1b-25b1aa9d075b
ms.date: 06/08/2017
---


# ChartGroup.HasHiLoLines Property (PowerPoint)

 **True** if the line chart has high-low lines. Read/write **Boolean**.


## Syntax

 _expression_. **HasHiLoLines**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

This property applies only to line charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables high-low lines for chart group one of the first chart in the active document and then sets line style, weight, and color. You should run the example on a 2-D line chart that has three series of stock-quote-like data (high-low-close).




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

