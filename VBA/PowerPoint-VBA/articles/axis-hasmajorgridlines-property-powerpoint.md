---
title: Axis.HasMajorGridlines Property (PowerPoint)
keywords: vbapp10.chm682008
f1_keywords:
- vbapp10.chm682008
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.HasMajorGridlines
ms.assetid: a8d5a060-ce84-8ca5-a42c-4a52d09a1e50
ms.date: 06/08/2017
---


# Axis.HasMajorGridlines Property (PowerPoint)

 **True** if the axis has major gridlines. Read/write **Boolean**.


## Syntax

 _expression_. **HasMajorGridlines**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Only axes in the primary axis group can have gridlines.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the color of the major gridlines for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            If .HasMajorGridlines Then

                ' Set the color to red.

                .MajorGridlines.Border.ColorIndex = 3

            End If

        End With

    End If

End With


```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

