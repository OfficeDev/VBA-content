---
title: Series.ApplyPictToSides Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.ApplyPictToSides
ms.assetid: b8a5b93d-f674-3927-3742-7578656f3152
ms.date: 06/08/2017
---


# Series.ApplyPictToSides Property (PowerPoint)

 **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean**.


## Syntax

 _expression_. **ApplyPictToSides**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example applies pictures to the sides of all points in the first series of the first chart in the active document. The series must already have pictures applied to it (setting this property changes the picture orientation).




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).ApplyPictToSides = True

    End If

End With


```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

