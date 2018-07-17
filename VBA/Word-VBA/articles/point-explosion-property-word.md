---
title: Point.Explosion Property (Word)
keywords: vbawd10.chm262144182
f1_keywords:
- vbawd10.chm262144182
ms.prod: word
api_name:
- Word.Point.Explosion
ms.assetid: e5305b4c-0ec5-79b3-6c71-2226fc3635ee
ms.date: 06/08/2017
---


# Point.Explosion Property (Word)

Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Read/write  **Long** .


## Syntax

 _expression_ . **Explosion**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Remarks

This property returns 0 (zero) if there is no explosion (the tip of the slice is in the center of the pie). 


## Example

The following example sets the explosion value for point two of the first chart in the active document. You should run the example on a pie chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Points(2).Explosion = 20 
 End If 
End With
```


## See also


#### Concepts


[Point Object](point-object-word.md)

