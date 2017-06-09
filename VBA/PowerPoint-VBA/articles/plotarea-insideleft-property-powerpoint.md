---
title: PlotArea.InsideLeft Property (PowerPoint)
keywords: vbapp10.chm67203
f1_keywords:
- vbapp10.chm67203
ms.prod: powerpoint
api_name:
- PowerPoint.PlotArea.InsideLeft
ms.assetid: 3357e9cd-4019-a8bd-48d3-d4f25348dd7b
ms.date: 06/08/2017
---


# PlotArea.InsideLeft Property (PowerPoint)

Returns or sets the distance, in points, from the chart edge to the inside left edge of the plot area. Read/write  **Double**.


## Syntax

 _expression_. **InsideLeft**

 _expression_ A variable that represents a **[PlotArea](plotarea-object-powerpoint.md)** object.


## Remarks

The plot area used for this measurement does not include the axis labels. The  **[Left](plotarea-left-property-powerpoint.md)** property for the plot area uses the bounding rectangle that includes the axis labels.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example draws a dotted rectangle around the inside of the plot area for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        With .Chart
            Set pa = .PlotArea
            With .Shapes.AddShape(msoShapeRectangle, _
                    pa.InsideLeft, pa.InsideTop, _
                    pa.InsideWidth, pa.InsideHeight)

                .Fill.Transparency = 1
                .Line.DashStyle = msoLineDashDot
            End With
        End With
    End If
End With
```


## See also


#### Concepts


[PlotArea Object](plotarea-object-powerpoint.md)

