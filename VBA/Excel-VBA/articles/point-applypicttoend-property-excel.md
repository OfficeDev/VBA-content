---
title: Point.ApplyPictToEnd Property (Excel)
keywords: vbaxl10.chm576096
f1_keywords:
- vbaxl10.chm576096
ms.prod: excel
api_name:
- Excel.Point.ApplyPictToEnd
ms.assetid: 9f814b2a-6c39-c0d9-0869-0df023c60e2c
ms.date: 06/08/2017
---


# Point.ApplyPictToEnd Property (Excel)

 **True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToEnd**

 _expression_ A variable that represents a **Point** object.


## Example

This example applies pictures to the end of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
Charts(1).SeriesCollection(1).ApplyPictToEnd = True
```


## See also


#### Concepts


[Point Object](point-object-excel.md)

