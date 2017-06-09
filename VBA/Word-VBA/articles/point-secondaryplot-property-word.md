---
title: Point.SecondaryPlot Property (Word)
keywords: vbawd10.chm262145662
f1_keywords:
- vbawd10.chm262145662
ms.prod: word
api_name:
- Word.Point.SecondaryPlot
ms.assetid: 89e56434-2b5a-b93c-cf18-8045cdf2da96
ms.date: 06/08/2017
---


# Point.SecondaryPlot Property (Word)

 **True** if the point is in the secondary section of either a pie-of-pie chart or a bar-of-pie chart. Read/write **Boolean** .


## Syntax

 _expression_ . **SecondaryPlot**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Remarks

This property applies only to points on pie-of-pie charts or bar-of-pie charts. 


## Example

The following example moves point four to the secondary section of the chart. You must run this example on either a pie-of-pie chart or a bar-of-pie chart. 


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1) 
 .Points(4).SecondaryPlot = True 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Point Object](point-object-word.md)

