---
title: Axis.CategoryNames Property (PowerPoint)
keywords: vbapp10.chm682004
f1_keywords:
- vbapp10.chm682004
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.CategoryNames
ms.assetid: f076ad9f-819b-4ced-a967-2fbda72fdfe8
ms.date: 06/08/2017
---


# Axis.CategoryNames Property (PowerPoint)

Returns or sets all the category names as a text array for the specified axis. Read/write  **Variant**.


## Syntax

 _expression_. **CategoryNames**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example uses an array to set individual category names for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Axes(xlCategory).CategoryNames = _
            Array ("1985", "1986", "1987", "1988", "1989")
    End If
End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

