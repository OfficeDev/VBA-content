---
title: Point.HasDataLabel Property (Excel)
keywords: vbaxl10.chm576081
f1_keywords:
- vbaxl10.chm576081
ms.prod: excel
api_name:
- Excel.Point.HasDataLabel
ms.assetid: 924f70a0-fdeb-e155-c857-55e0dfb7ca60
ms.date: 06/08/2017
---


# Point.HasDataLabel Property (Excel)

 **True** if the point has a data label. Read/write **Boolean** .


## Syntax

 _expression_ . **HasDataLabel**

 _expression_ A variable that represents a **Point** object.


## Example

This example turns on the data label for point seven in series three in Chart1, and then it sets the data label color to blue.


```vb
With Charts("Chart1").SeriesCollection(3).Points(7) 
 .HasDataLabel = True 
 .ApplyDataLabels Type:=xlValue 
 .DataLabel.Font.ColorIndex = 5 
End With
```


## See also


#### Concepts


[Point Object](point-object-excel.md)

