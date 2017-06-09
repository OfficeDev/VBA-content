---
title: Series.ApplyPictToSides Property (Excel)
keywords: vbaxl10.chm578115
f1_keywords:
- vbaxl10.chm578115
ms.prod: excel
api_name:
- Excel.Series.ApplyPictToSides
ms.assetid: 300e9c75-4108-32bc-01ab-c622843e6fbd
ms.date: 06/08/2017
---


# Series.ApplyPictToSides Property (Excel)

 **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyPictToSides**

 _expression_ A variable that represents a **Series** object.


## Example

This example applies pictures to the sides of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
Charts(1).SeriesCollection(1).ApplyPictToSides = True
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

