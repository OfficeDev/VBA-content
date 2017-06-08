---
title: Chart.BarShape Property (PowerPoint)
keywords: vbapp10.chm684005
f1_keywords:
- vbapp10.chm684005
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.BarShape
ms.assetid: fae46b36-9d4c-3646-db91-95691d8ce2af
ms.date: 06/08/2017
---


# Chart.BarShape Property (PowerPoint)

Returns or sets the shape used for every series in a 3-D bar or column chart. Read/write  **[XlBarShape](xlbarshape-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **BarShape**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the shape used with the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.BarShape = xlConeToPoint

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

