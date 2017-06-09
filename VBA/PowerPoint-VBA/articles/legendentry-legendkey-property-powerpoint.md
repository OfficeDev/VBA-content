---
title: LegendEntry.LegendKey Property (PowerPoint)
keywords: vbapp10.chm65710
f1_keywords:
- vbapp10.chm65710
ms.prod: powerpoint
api_name:
- PowerPoint.LegendEntry.LegendKey
ms.assetid: 6265569c-fc7c-5fe8-864e-d543a08b33f4
ms.date: 06/08/2017
---


# LegendEntry.LegendKey Property (PowerPoint)

Returns the legend key that is associated with the entry. Read-only  **[LegendKey](legendkey-object-powerpoint.md)**.


## Syntax

 _expression_. **LegendKey**

 _expression_ A variable that represents a **[LegendEntry](legendentry-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the legend key for legend entry one on the first chart in the active document to be a triangle. You should run the example on a 2-D line chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Legend.LegendEntries(1).LegendKey _
            .MarkerStyle = xlMarkerStyleTriangle
    End If
End With
```


## See also


#### Concepts


[LegendEntry Object](legendentry-object-powerpoint.md)

