---
title: Axis.MinimumScaleIsAuto Property (PowerPoint)
keywords: vbapp10.chm682020
f1_keywords:
- vbapp10.chm682020
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MinimumScaleIsAuto
ms.assetid: 7ec5b07d-3683-e45b-ca39-d67ce959edfc
ms.date: 06/08/2017
---


# Axis.MinimumScaleIsAuto Property (PowerPoint)

 **True** if Microsoft Word calculates the minimum value for the value axis. Read/write **Boolean**.


## Syntax

 _expression_. **MinimumScaleIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting the  **[MinimumScale](axis-minimumscale-property-powerpoint.md)** property sets this property to **False**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example automatically calculates the minimum scale and the maximum scale for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .MinimumScaleIsAuto = True

            .MaximumScaleIsAuto = True

        End With

    End If

End With


```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

