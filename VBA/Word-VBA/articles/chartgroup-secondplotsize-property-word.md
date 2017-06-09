---
title: ChartGroup.SecondPlotSize Property (Word)
keywords: vbawd10.chm263454764
f1_keywords:
- vbawd10.chm263454764
ms.prod: word
api_name:
- Word.ChartGroup.SecondPlotSize
ms.assetid: 68f4d170-62c8-eb34-26a2-693aa96fc5f1
ms.date: 06/08/2017
---


# ChartGroup.SecondPlotSize Property (Word)

Returns or sets the size, as a percentage of the primary pie, of the secondary section of either a pie-of-pie chart or a bar-of-pie chart. Read/write  **Long** .


## Syntax

 _expression_ . **SecondPlotSize**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

This property can have a value from 5 through 200. 


## Example

The following example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. The secondary section is 50 percent of the size of the primary pie. You must run the example on either a pie-of-pie chart or a bar-of-pie chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
 .SecondPlotSize = 50 
 End With 
 End If 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

