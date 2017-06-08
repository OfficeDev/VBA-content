---
title: Axis.CategoryType Property (PowerPoint)
keywords: vbapp10.chm682037
f1_keywords:
- vbapp10.chm682037
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.CategoryType
ms.assetid: bbcb485d-9464-33c8-ca9b-e3463bc9e884
ms.date: 06/08/2017
---


# Axis.CategoryType Property (PowerPoint)

Returns or sets the category axis type. Read/write  **[XlCategoryType](xlcategorytype-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **CategoryType**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

You cannot set this property for a value axis.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis for the first chart in the active document to use a time scale, using months as the base unit.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .CategoryType = xlTimeScale

            .BaseUnit = xlMonths

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

