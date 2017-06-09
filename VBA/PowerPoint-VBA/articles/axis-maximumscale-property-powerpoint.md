---
title: Axis.MaximumScale Property (PowerPoint)
keywords: vbapp10.chm682017
f1_keywords:
- vbapp10.chm682017
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MaximumScale
ms.assetid: cb0588ce-0685-77ac-da06-75a913f90e41
ms.date: 06/08/2017
---


# Axis.MaximumScale Property (PowerPoint)

Returns or sets the maximum value on the value axis. Read/write  **Double**.


## Syntax

 _expression_. **MaximumScale**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting this property sets the  **[MaximumScaleIsAuto](axis-maximumscaleisauto-property-powerpoint.md)** property to **False**.


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

