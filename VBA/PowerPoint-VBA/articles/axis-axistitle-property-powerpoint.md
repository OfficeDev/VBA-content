---
title: Axis.AxisTitle Property (PowerPoint)
keywords: vbapp10.chm682003
f1_keywords:
- vbapp10.chm682003
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.AxisTitle
ms.assetid: c1063cf8-3aa2-ea39-ea2d-33a7c63b77d4
ms.date: 06/08/2017
---


# Axis.AxisTitle Property (PowerPoint)

Returns the title of the specified axis. Read-only  **[AxisTitle](axistitle-object-powerpoint.md)**.


## Syntax

 _expression_. **AxisTitle**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds an axis label to the category axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .HasTitle = True

            .AxisTitle.Text = "July Sales"

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

