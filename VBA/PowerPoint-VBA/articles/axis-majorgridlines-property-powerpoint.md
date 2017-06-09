---
title: Axis.MajorGridlines Property (PowerPoint)
keywords: vbapp10.chm682011
f1_keywords:
- vbapp10.chm682011
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MajorGridlines
ms.assetid: d0ec2384-8503-0198-388c-c74231137bf0
ms.date: 06/08/2017
---


# Axis.MajorGridlines Property (PowerPoint)

Returns the major gridlines for the specified axis. Read-only  **[Gridlines](gridlines-object-powerpoint.md)**.


## Syntax

 _expression_. **MajorGridlines**

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

                 ' Set the color to blue.

                .MajorGridlines.Border.ColorIndex = 5 

            End If

        End With

    End If

End With


```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

