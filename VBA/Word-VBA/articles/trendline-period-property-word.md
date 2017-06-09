---
title: Trendline.Period Property (Word)
keywords: vbawd10.chm26345656
f1_keywords:
- vbawd10.chm26345656
ms.prod: word
api_name:
- Word.Trendline.Period
ms.assetid: 3e9a1a9d-4f88-1ba2-d2c7-ed4d1b6ec514
ms.date: 06/08/2017
---


# Trendline.Period Property (Word)

Returns or sets the period for the moving-average trendline. Read/write  **Long** .


## Syntax

 _expression_ . **Period**

 _expression_ A variable that represents a **[Trendline](trendline-object-word.md)** object.


## Remarks

This property can have a value from 2 through 255. 


## Example

The following example sets the period for the moving-average trendline on the first chart in the active document. You should run the example on a 2-D column chart that has a single series that contains 10 data points and a moving-average trendline.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Trendlines(1) 
 If .Type = xlMovingAvg Then .Period = 5 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Trendline Object](trendline-object-word.md)

