---
title: Axis.MinorUnitScale Property (PowerPoint)
keywords: vbapp10.chm682036
f1_keywords:
- vbapp10.chm682036
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MinorUnitScale
ms.assetid: 15ce78c6-b054-afea-bd6c-6a40db7f93aa
ms.date: 06/08/2017
---


# Axis.MinorUnitScale Property (PowerPoint)

Returns or sets the minor unit scale value for the category axis when the  **[CategoryType](axis-categorytype-property-powerpoint.md)** property is set to **xlTimeScale**. Read/write **[XlTimeUnit](xltimeunit-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **MinorUnitScale**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

 **MinorUnitScale** can be one of the following **XlTimeUnit** constants:


-  **xlMonths**
    
-  **xlDays**
    
-  **xlYears**
    

## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis to use a time scale and sets the major and minor units.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .CategoryType = xlTimeScale

            .MajorUnit = 5

            .MajorUnitScale = xlDays

            .MinorUnit = 1

            .MinorUnitScale = xlDays

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

