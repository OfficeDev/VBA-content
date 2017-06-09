---
title: Series.ApplyPictToFront Property (Excel)
keywords: vbaxl10.chm578116
f1_keywords:
- vbaxl10.chm578116
ms.prod: excel
api_name:
- Excel.Series.ApplyPictToFront
ms.assetid: b40a8808-734f-a00e-3fa1-4e1cf5ac0a52
ms.date: 06/08/2017
---


# Series.ApplyPictToFront Property (Excel)

 **True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToFront**

 _expression_ A variable that represents a **Series** object.


## Example

This example applies pictures to the front of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
Charts(1).SeriesCollection(1).ApplyPictToFront = True
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

