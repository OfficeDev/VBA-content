---
title: DownBars Object (PowerPoint)
keywords: vbapp10.chm700000
f1_keywords:
- vbapp10.chm700000
ms.prod: powerpoint
api_name:
- PowerPoint.DownBars
ms.assetid: ce479049-2e58-2dad-f4bb-2dd27a223753
ms.date: 06/08/2017
---


# DownBars Object (PowerPoint)

Represents the down bars in a chart group.


## Remarks

 Down bars connect points on the first series in the chart group with lower values on the last series (the lines go down from the first series). Only 2-D line groups that contain at least two series can have down bars. This object is not a collection. There is no object that represents a single down bar; you either enable up bars and down bars for all points in a chart group or you disable them.

If the  **[HasUpDownBars](chartgroup-hasupdownbars-property-powerpoint.md)** property is **False**, most properties of the **DownBars** object are disabled.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[DownBars](chartgroup-downbars-property-powerpoint.md)** property to return the **DownBars** object. The following example enables up and down bars for chart group one of the first chart in the active document. The example then sets the up bar color to blue and the down bar color to red.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasUpDownBars = True

            .UpBars.Interior.Color = RGB(0, 0, 255)

            .DownBars.Interior.Color = RGB(255, 0, 0)

        End With

    End If

End With


```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)


