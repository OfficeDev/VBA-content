---
title: ChartFont.Color Property (PowerPoint)
keywords: vbapp10.chm704003
f1_keywords:
- vbapp10.chm704003
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.Color
ms.assetid: c95350e1-11c3-f105-15d8-9d2686cc1340
ms.date: 06/08/2017
---


# ChartFont.Color Property (PowerPoint)

Returns or sets the primary color of the object. Read/write  **Variant**.


## Syntax

 _expression_. **Color**

 _expression_ A variable that represents a **[ChartFont](chartfont-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the color of the tick-mark labels on the value axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        Chart.Axes(xlValue).TickLabels.Font.Color = _
            RGB(0, 255, 0)
    End If
End With
```


## See also


#### Concepts


[ChartFont Object](chartfont-object-powerpoint.md)

