---
title: Axis.DisplayUnit Property (PowerPoint)
keywords: vbapp10.chm682042
f1_keywords:
- vbapp10.chm682042
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.DisplayUnit
ms.assetid: 6545b191-ef58-49d5-2df3-04d0d0d06476
ms.date: 06/08/2017
---


# Axis.DisplayUnit Property (PowerPoint)

Returns or sets the unit label for the value axis. Read/write  **[XlDisplayUnit](xldisplayunit-enumeration-powerpoint.md)**, **xlCustom**, or **xlNone**.


## Syntax

 _expression_. **DisplayUnit**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Using unit labels when charting large values makes your tick-mark labels easier to read. For example, if you label your value axis in units of hundreds, thousands, or millions, you can use smaller numeric values at the tick marks on the axis.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the units displayed on the value axis of the first chart in the active document to hundreds.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .DisplayUnit = xlHundreds

            .HasTitle = True

            .AxisTitle.Caption = "Rebate Amounts"

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

