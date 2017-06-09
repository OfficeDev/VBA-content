---
title: Axis.MinorGridlines Property (PowerPoint)
keywords: vbapp10.chm682021
f1_keywords:
- vbapp10.chm682021
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MinorGridlines
ms.assetid: f9e1168d-af71-6876-a289-a9e8d1db38cb
ms.date: 06/08/2017
---


# Axis.MinorGridlines Property (PowerPoint)

Returns the minor gridlines for the specified axis. Read-only  **[Gridlines](gridlines-object-powerpoint.md)**.


## Syntax

 _expression_. **MinorGridlines**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Only axes in the primary axis group can have gridlines.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the color of the minor gridlines for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            If .HasMinorGridlines Then

                ' Set the color to blue.

                .MinorGridlines.Border.ColorIndex = 5

            End If

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

