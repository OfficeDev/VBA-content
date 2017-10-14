---
title: Axis.MaximumScaleIsAuto Property (PowerPoint)
keywords: vbapp10.chm682018
f1_keywords:
- vbapp10.chm682018
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MaximumScaleIsAuto
ms.assetid: f25fd6a9-4ca7-2f06-3db4-35002f1c91ae
ms.date: 06/08/2017
---


# Axis.MaximumScaleIsAuto Property (PowerPoint)

 **True** if Microsoft Word calculates the maximum value for the value axis. Read/write **Boolean**.


## Syntax

 _expression_. **MaximumScaleIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting the  **[MaximumScale](axis-maximumscale-property-powerpoint.md)** property sets this property to **False**.


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

