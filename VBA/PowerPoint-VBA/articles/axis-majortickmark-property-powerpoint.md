---
title: Axis.MajorTickMark Property (PowerPoint)
keywords: vbapp10.chm682012
f1_keywords:
- vbapp10.chm682012
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MajorTickMark
ms.assetid: 82397f1c-8a0d-44dd-a9f3-3426fee03f1d
ms.date: 06/08/2017
---


# Axis.MajorTickMark Property (PowerPoint)

Returns or sets the type of major tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **MajorTickMark**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

 **MajorTickMark** can be set to one of the following **XlTickMark** constants:


-  **xlTickMarkInside**
    
-  **xlTickMarkOutside**
    
-  **xlTickMarkCross**
    
-  **xlTickMarkNone**
    

## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the major tick marks for the value axis for the first chart in the active document to be outside the axis.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlValue).MajorTickMark = xlTickMarkOutside

    End If

End With


```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

