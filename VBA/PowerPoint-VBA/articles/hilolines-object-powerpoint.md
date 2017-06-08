---
title: HiLoLines Object (PowerPoint)
keywords: vbapp10.chm706000
f1_keywords:
- vbapp10.chm706000
ms.prod: powerpoint
api_name:
- PowerPoint.HiLoLines
ms.assetid: 77a7ae91-daf3-4c35-1f39-067d2698fb43
ms.date: 06/08/2017
---


# HiLoLines Object (PowerPoint)

Represents the high-low lines in a chart group.


## Remarks

 High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2-D line groups can have high-low lines. This object is not a collection. There is no object that represents a single high-low line; you either enable high-low lines for all points in a chart group or disable them.

If the  **[HasHiLoLines](chartgroup-hashilolines-property-powerpoint.md)** property is **False**, most properties of the **HiLoLines** object are disabled.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[HiLoLines](chartgroup-hilolines-property-powerpoint.md)** property to return the **HiLoLines** object. The following example uses the **HasHiLowLines** property to add high-low lines to the first chart (the chart must be a line chart) in the active document. The example then makes the high-low lines blue.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With Chart.ChartGroups(1)

            .HasHighLowLines = True

            .HiLoLines.Border.Color = RGB(0, 0, 255)

        End With

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

