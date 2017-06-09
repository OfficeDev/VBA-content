---
title: Point.DataLabel Property (Excel)
keywords: vbaxl10.chm576078
f1_keywords:
- vbaxl10.chm576078
ms.prod: excel
api_name:
- Excel.Point.DataLabel
ms.assetid: 2f860d46-c6b5-50cf-b0af-4c46d9f7b2ac
ms.date: 06/08/2017
---


# Point.DataLabel Property (Excel)

Returns a  **[DataLabel](datalabel-object-excel.md)** object that represents the data label associated with the point. Read-only.


## Syntax

 _expression_ . **DataLabel**

 _expression_ A variable that represents a **Point** object.


## Example

This example turns on the data label for point seven in series three in Chart1, and then it sets the data label color to blue.


```vb
With Charts("Chart1").SeriesCollection(3).Points(7) 
 .HasDataLabel = True 
 .ApplyDataLabels type:=xlValue 
 .DataLabel.Font.ColorIndex = 5 
End With
```


## See also


#### Concepts


[Point Object](point-object-excel.md)

