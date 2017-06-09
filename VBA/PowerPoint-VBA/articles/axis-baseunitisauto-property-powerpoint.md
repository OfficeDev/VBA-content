---
title: Axis.BaseUnitIsAuto Property (PowerPoint)
keywords: vbapp10.chm682034
f1_keywords:
- vbapp10.chm682034
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.BaseUnitIsAuto
ms.assetid: 3cc90d1a-a87f-ac57-b2a2-bf3ccc964a8e
ms.date: 06/08/2017
---


# Axis.BaseUnitIsAuto Property (PowerPoint)

 **True** if Microsoft Word chooses appropriate base units for the specified category axis. The default is **True**. Read/write **Boolean**.


## Syntax

 _expression_. **BaseUnitIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

You cannot set this property for a value axis.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis for the first chart in the active document to use a time scale, with the base unit automatically chosen by Word.




```


With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .Axes(xlCategory).CategoryType = xlTimeScale

            .Axes(xlCategory).BaseUnitIsAuto = True

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

