---
title: ChartGroup.FirstSliceAngle Property (PowerPoint)
keywords: vbapp10.chm692010
f1_keywords:
- vbapp10.chm692010
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.FirstSliceAngle
ms.assetid: fb09ab99-9a85-3932-f569-56b5bbb87b50
ms.date: 06/08/2017
---


# ChartGroup.FirstSliceAngle Property (PowerPoint)

Returns or sets the angle, in degrees (clockwise from vertical), of the first pie-chart or doughnut-chart slice. Read/write  **Long**.


## Syntax

 _expression_. **FirstSliceAngle**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

This property applies only to pie, 3-D pie, and doughnut charts. It can be a value from 0 through 360. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the angle for the first slice in chart group one for the first chart in the active document. You should run the example on a 2-D pie chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups(1).FirstSliceAngle = 15

    End If

End With


```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

