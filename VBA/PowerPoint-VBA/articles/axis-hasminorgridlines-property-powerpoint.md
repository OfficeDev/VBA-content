---
title: Axis.HasMinorGridlines Property (PowerPoint)
keywords: vbapp10.chm682009
f1_keywords:
- vbapp10.chm682009
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.HasMinorGridlines
ms.assetid: 4ee1c716-296b-eeaf-8d14-bcb6e0919611
ms.date: 06/08/2017
---


# Axis.HasMinorGridlines Property (PowerPoint)

 **True** if the axis has minor gridlines. Read/write **Boolean**.


## Syntax

 _expression_. **HasMinorGridlines**

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

                ' Set the color to green.

                .MinorGridlines.Border.ColorIndex = 4

            End If

        End With

    End If

End With


```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

