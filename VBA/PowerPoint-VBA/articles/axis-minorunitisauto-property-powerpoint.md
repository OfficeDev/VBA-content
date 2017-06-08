---
title: Axis.MinorUnitIsAuto Property (PowerPoint)
keywords: vbapp10.chm682024
f1_keywords:
- vbapp10.chm682024
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MinorUnitIsAuto
ms.assetid: 18dff25c-59a3-e2c8-2997-6239b1ae87bf
ms.date: 06/08/2017
---


# Axis.MinorUnitIsAuto Property (PowerPoint)

 **True** if Microsoft Word calculates minor units for the value axis. Read/write **Boolean**.


## Syntax

 _expression_. **MinorUnitIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting the  **[MinorUnit](axis-minorunit-property-powerpoint.md)** property sets this property to **False**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example automatically calculates major and minor units for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .MajorUnitIsAuto = True

            .MinorUnitIsAuto = True

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

