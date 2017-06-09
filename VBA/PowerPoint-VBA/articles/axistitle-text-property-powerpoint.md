---
title: AxisTitle.Text Property (PowerPoint)
keywords: vbapp10.chm683008
f1_keywords:
- vbapp10.chm683008
ms.prod: powerpoint
api_name:
- PowerPoint.AxisTitle.Text
ms.assetid: c498054e-1b96-66c2-e4c3-d06c72935552
ms.date: 06/08/2017
---


# AxisTitle.Text Property (PowerPoint)

Returns or sets the text for the specified object. Read/write  **String**.


## Syntax

 _expression_. **Text**

 _expression_ A variable that represents an **[AxisTitle](axistitle-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the axis title text for the category axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .HasTitle = True

            .AxisTitle.Text = "Month"

        End With

    End If

End With
```


## See also


#### Concepts


[AxisTitle Object](axistitle-object-powerpoint.md)

