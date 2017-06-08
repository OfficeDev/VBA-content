---
title: Axis.MinimumScale Property (PowerPoint)
keywords: vbapp10.chm682019
f1_keywords:
- vbapp10.chm682019
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MinimumScale
ms.assetid: 90cfaa99-0ccf-2fa8-c6b0-54d1440b6677
ms.date: 06/08/2017
---


# Axis.MinimumScale Property (PowerPoint)

Returns or sets the minimum value on the value axis. Read/write  **Double**.


## Syntax

 _expression_. **MinimumScale**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting this property sets the  **[MinimumScaleIsAuto](axis-minimumscaleisauto-property-powerpoint.md)** property to **False**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the minimum and maximum values for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .MinimumScale = 10

            .MaximumScale = 120

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

