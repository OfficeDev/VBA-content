---
title: Axis.BaseUnit Property (PowerPoint)
keywords: vbapp10.chm682033
f1_keywords:
- vbapp10.chm682033
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.BaseUnit
ms.assetid: a53e90c5-5048-8e93-57b2-024d64d2ff73
ms.date: 06/08/2017
---


# Axis.BaseUnit Property (PowerPoint)

Returns or sets the base unit for the specified category axis. Read/write  **[XlTimeUnit](xltimeunit-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **BaseUnit**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting this property has no visible effect if the  **[CategoryType](axis-categorytype-property-powerpoint.md)** property for the specified axis is set to **xlCategoryScale**. The set value is retained, however, and takes effect when the **CategoryType** property is set to **xlTimeScale**.

You cannot set this property for a value axis.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis for the first chart in the active document to use a time scale, using months as the base unit.




```


With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .Axes(xlCategory).CategoryType = xlTimeScale

            .Axes(xlCategory).BaseUnit = xlMonths

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

