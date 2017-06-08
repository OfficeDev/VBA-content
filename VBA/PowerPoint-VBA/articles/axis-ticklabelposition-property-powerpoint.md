---
title: Axis.TickLabelPosition Property (PowerPoint)
keywords: vbapp10.chm682028
f1_keywords:
- vbapp10.chm682028
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.TickLabelPosition
ms.assetid: 439b3da0-37d1-1fd8-b810-66accac03001
ms.date: 06/08/2017
---


# Axis.TickLabelPosition Property (PowerPoint)

Describes the position of tick-mark labels on the specified axis. Read/write  **[XlTickLabelPosition](xlticklabelposition-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **TickLabelPosition**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets tick-mark labels to the high position (above the chart) on the category axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Axes(xlCategory) _
            .TickLabelPosition = xlTickLabelPositionHigh
    End If
End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

