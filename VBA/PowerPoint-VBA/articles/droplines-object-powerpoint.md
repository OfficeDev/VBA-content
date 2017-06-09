---
title: DropLines Object (PowerPoint)
keywords: vbapp10.chm701000
f1_keywords:
- vbapp10.chm701000
ms.prod: powerpoint
api_name:
- PowerPoint.DropLines
ms.assetid: b13b58c3-d00d-16d2-16ef-bcd3cae347c5
ms.date: 06/08/2017
---


# DropLines Object (PowerPoint)

Represents the drop lines in a chart group.


## Remarks

Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines. This object is not a collection. There is no object that represents a single drop line; you either enable drop lines for all points in a chart group or you disable them.

If the  **[HasDropLines](chartgroup-hasdroplines-property-powerpoint.md)** property is **False**, most properties of the **DropLines** object are disabled.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[DropLines](chartgroup-droplines-property-powerpoint.md)** property to return the **DropLines** object. The following example enables drop lines for chart group one of the first chart in the active document and then sets the drop line color to red.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasDropLines = True

            .DropLines.Border.ColorIndex = 3

        End With

    End If

End With
```


## See also


#### Concepts



[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

