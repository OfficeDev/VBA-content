---
title: Trendline.DisplayRSquared Property (Word)
keywords: vbawd10.chm26345661
f1_keywords:
- vbawd10.chm26345661
ms.prod: word
api_name:
- Word.Trendline.DisplayRSquared
ms.assetid: 10f55d97-f9f2-953a-427b-b158abe268d7
ms.date: 06/08/2017
---


# Trendline.DisplayRSquared Property (Word)

 **True** if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayRSquared**

 _expression_ A variable that represents a **[Trendline](trendline-object-word.md)** object.


## Remarks

Setting this property to  **True** automatically turns on data labels.


## Example

The following example displays the R-squared value and equation for trendline one of the first chart in the active document. You should run the example on a 2-D column chart that has a trendline for the first series.


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

