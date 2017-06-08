---
title: UpBars Object (PowerPoint)
keywords: vbapp10.chm722000
f1_keywords:
- vbapp10.chm722000
ms.prod: powerpoint
api_name:
- PowerPoint.UpBars
ms.assetid: 8a176f01-01a6-86bc-a69b-29763ebb1481
ms.date: 06/08/2017
---


# UpBars Object (PowerPoint)

Represents the up bars in a chart group.


## Remarks

Up bars connect points on series one with higher values on the last series in the chart group (the lines go up from series one). Only 2-D line groups that contain at least two series can have up bars. This object is not a collection. There is no object that represents a single up bar; you either enable up bars for all points in a chart group or you disable them.

If the  **[HasUpDownBars](chartgroup-hasupdownbars-property-powerpoint.md)** property is **False**, most properties of the **UpBars** object are disabled.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[UpBars](chartgroup-upbars-property-powerpoint.md)** property to return the **UpBars** object. The following example enables up and down bars for chart group one of the first chart in the active document. The example then sets the up bar color to blue and sets the down bar color to red.




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

