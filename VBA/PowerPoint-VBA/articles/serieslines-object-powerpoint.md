---
title: SeriesLines Object (PowerPoint)
keywords: vbapp10.chm718000
f1_keywords:
- vbapp10.chm718000
ms.prod: powerpoint
api_name:
- PowerPoint.SeriesLines
ms.assetid: 5d953ed4-ca16-3cb3-ba8f-1742e4a56cb6
ms.date: 06/08/2017
---


# SeriesLines Object (PowerPoint)

Represents series lines in a chart group.


## Remarks

 Series lines connect the data values from each series. Only 2-D stacked bar, 2-D stacked column, pie-of-pie, or bar-of-pie charts can have series lines. This object is not a collection. There is no object that represents a single series line; you either enable series lines for all points in a chart group or you disable them.

If the  **[HasSeriesLines](chartgroup-hasserieslines-property-powerpoint.md)** property is **False**, most properties of the **SeriesLines** object are disabled.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[SeriesLines](chartgroup-serieslines-property-powerpoint.md)** property to return a **SeriesLines** object. The following example adds series lines to chart group one in embedded chart one on worksheet one (the chart must be a 2-D stacked bar or column chart).




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasSeriesLines = True

            .SeriesLines.Border.Color = RGB(0, 0, 255)

        End With

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

