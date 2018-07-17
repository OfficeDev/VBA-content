---
title: Axis.MajorUnitIsAuto Property (PowerPoint)
keywords: vbapp10.chm682016
f1_keywords:
- vbapp10.chm682016
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MajorUnitIsAuto
ms.assetid: ffea2f83-1a5e-7ae1-f866-ae52a4d49567
ms.date: 06/08/2017
---


# Axis.MajorUnitIsAuto Property (PowerPoint)

 **True** if Microsoft Word calculates the major units for the value axis. Read/write **Boolean**.


## Syntax

 _expression_. **MajorUnitIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting the  **[MajorUnit](axis-majorunit-property-powerpoint.md)** property sets this property to **False**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example automatically sets the major and minor units for the value axis of the first chart in the active document.




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

