---
title: Series.BarShape Property (PowerPoint)
keywords: vbapp10.chm66939
f1_keywords:
- vbapp10.chm66939
ms.prod: powerpoint
api_name:
- PowerPoint.Series.BarShape
ms.assetid: c6f2d0b7-407e-4073-89b1-433e9386287a
ms.date: 06/08/2017
---


# Series.BarShape Property (PowerPoint)

Returns or sets the shape used for a single series in a 3-D bar or column chart. Read/write  **[XlBarShape](xlbarshape-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **BarShape**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the shape used for the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).BarShape = xlConeToPoint

    End If

End With
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

