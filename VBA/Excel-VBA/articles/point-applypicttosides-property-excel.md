---
title: Point.ApplyPictToSides Property (Excel)
keywords: vbaxl10.chm576094
f1_keywords:
- vbaxl10.chm576094
ms.prod: excel
api_name:
- Excel.Point.ApplyPictToSides
ms.assetid: 46513ac1-9a83-a6cf-ef09-f5075b2df66f
ms.date: 06/08/2017
---


# Point.ApplyPictToSides Property (Excel)

 **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToSides**

 _expression_ A variable that represents a **Point** object.


## Example

This example applies pictures to the sides of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
Charts(1).SeriesCollection(1).ApplyPictToSides = True
```


## See also


#### Concepts


[Point Object](point-object-excel.md)

