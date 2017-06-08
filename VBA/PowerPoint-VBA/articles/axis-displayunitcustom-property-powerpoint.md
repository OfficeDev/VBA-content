---
title: Axis.DisplayUnitCustom Property (PowerPoint)
keywords: vbapp10.chm682043
f1_keywords:
- vbapp10.chm682043
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.DisplayUnitCustom
ms.assetid: bfee899d-27fd-ca15-9af7-04702ae3da52
ms.date: 06/08/2017
---


# Axis.DisplayUnitCustom Property (PowerPoint)

If the value of the  **[DisplayUnit](axis-displayunit-property-powerpoint.md)** property is **xlCustom**, returns or sets the value of the displayed units. Read/write **Double**.


## Syntax

 _expression_. **DisplayUnitCustom**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

The value of this property must be from 0 through 10E307.

Using unit labels when charting large values makes your tick-mark labels easier to read. For example, if you label your value axis in units of hundreds, thousands, or millions, you can use smaller numeric values at the tick marks on the axis.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the units displayed on the value axis of the first chart in the active document to increments of 500.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .DisplayUnit = xlCustom

            .DisplayUnitCustom = 500

            .HasTitle = True

            .AxisTitle.Caption = "Rebate Amounts"

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

