---
title: PlotArea.InsideTop Property (PowerPoint)
keywords: vbapp10.chm67204
f1_keywords:
- vbapp10.chm67204
ms.prod: powerpoint
api_name:
- PowerPoint.PlotArea.InsideTop
ms.assetid: 03f9c821-80f1-26db-580b-6e2e29e0ae9c
ms.date: 06/08/2017
---


# PlotArea.InsideTop Property (PowerPoint)

Returns or sets the distance, in points, from the chart edge to the inside top edge of the plot area. Read/write  **Double**.


## Syntax

 _expression_. **InsideTop**

 _expression_ A variable that represents a **[PlotArea](plotarea-object-powerpoint.md)** object.


## Remarks

The plot area used for this measurement does not include the axis labels. The  **[Top](plotarea-top-property-powerpoint.md)** property for the plot area uses the bounding rectangle that includes the axis labels.


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

