---
title: ChartGroup.HasSeriesLines Property (Word)
keywords: vbawd10.chm263454736
f1_keywords:
- vbawd10.chm263454736
ms.prod: word
api_name:
- Word.ChartGroup.HasSeriesLines
ms.assetid: 56e85d95-4743-4afd-5bdf-d00065608708
ms.date: 06/08/2017
---


# ChartGroup.HasSeriesLines Property (Word)

 **True** if a stacked column chart or bar chart has series lines or if a pie-of-pie chart or bar-of-pie chart has connector lines between the two sections. Read/write **Boolean** .


## Syntax

 _expression_ . **HasSeriesLines**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

This property applies only to 2-D stacked bar, 2-D stacked column, pie-of-pie, or bar-of-pie charts. 


## Example

The following example enables series lines for chart group one of the first chart in the active document and then sets their line style, weight, and color. You should run the example on a 2-D stacked column chart that has two or more series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasSeriesLines = True 
 With .SeriesLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

