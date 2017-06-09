---
title: Trendline.Backward2 Property (Word)
keywords: vbawd10.chm26348122
f1_keywords:
- vbawd10.chm26348122
ms.prod: word
api_name:
- Word.Trendline.Backward2
ms.assetid: 41655a35-671d-363b-e980-589a1d53ffe1
ms.date: 06/08/2017
---


# Trendline.Backward2 Property (Word)

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write  **Double** .


## Syntax

 _expression_ . **Backward2**

 _expression_ A variable that represents a **[Trendline](trendline-object-word.md)** object.


## Example

The following example sets the number of units that the trendline for the first chart in the active document extends forward and backward. You should run the example on a 2-D column chart that contains a single series that has a trendline.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Trendlines(1) 
 .Forward2 = 5 
 .Backward2 = .5 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Trendline Object](trendline-object-word.md)

