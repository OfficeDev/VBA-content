---
title: Trendline.DisplayEquation Property (Word)
keywords: vbawd10.chm26345662
f1_keywords:
- vbawd10.chm26345662
ms.prod: word
api_name:
- Word.Trendline.DisplayEquation
ms.assetid: c5534224-f7ff-2899-0d45-2c9fca8afbd8
ms.date: 06/08/2017
---


# Trendline.DisplayEquation Property (Word)

 **True** if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayEquation**

 _expression_ A variable that represents a **[Trendline](trendline-object-word.md)** object.


## Remarks

Setting this property to  **True** automatically enables data labels.


## Example

The following example displays the R-squared value and equation for the first trendline of the first chart in the active document. You should run the example on a 2-D column chart that has a trendline for the first series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = True 
 .DisplayEquation = True 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Trendline Object](trendline-object-word.md)

