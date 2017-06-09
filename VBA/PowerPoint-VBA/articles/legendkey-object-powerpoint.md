---
title: LegendKey Object (PowerPoint)
keywords: vbapp10.chm712000
f1_keywords:
- vbapp10.chm712000
ms.prod: powerpoint
api_name:
- PowerPoint.LegendKey
ms.assetid: 98e8b9c3-b53e-9595-9389-6f92a6d730f4
ms.date: 06/08/2017
---


# LegendKey Object (PowerPoint)

Represents a legend key in a chart legend.


## Remarks

 Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart. The legend key is linked to its associated series or trendline in such a way that changing the formatting of one simultaneously changes the formatting of the other.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[LegendKey](legendentry-legendkey-property-powerpoint.md)** property to return the **LegendKey** object. The following example changes the marker background color for the legend entry at the top of the legend for the first chart in the active document. This simultaneously changes the format of every point in the series associated with this legend entry. The associated series must support data markers.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Legend.LegendEntries(1).LegendKey _
            .MarkerBackgroundColorIndex = 5
    End If
End With


```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

