---
title: Point.ApplyPictToFront Property (Excel)
keywords: vbaxl10.chm576095
f1_keywords:
- vbaxl10.chm576095
ms.prod: excel
api_name:
- Excel.Point.ApplyPictToFront
ms.assetid: e739e368-9789-be23-da90-17ab4cf3a935
ms.date: 06/08/2017
---


# Point.ApplyPictToFront Property (Excel)

 **True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToFront**

 _expression_ A variable that represents a **Point** object.


## Example

This example applies pictures to the front of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
Charts(1).SeriesCollection(1).ApplyPictToFront = True
```


## See also


#### Concepts


[Point Object](point-object-excel.md)

