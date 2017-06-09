---
title: ChartGroup.Overlap Property (Word)
keywords: vbawd10.chm263454742
f1_keywords:
- vbawd10.chm263454742
ms.prod: word
api_name:
- Word.ChartGroup.Overlap
ms.assetid: e2d219f6-7edd-69c7-015f-8304cf95dbc3
ms.date: 06/08/2017
---


# ChartGroup.Overlap Property (Word)

Specifies how bars and columns are positioned. Read/write  **Long** .


## Syntax

 _expression_ . **Overlap**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

 This property applies only to 2-D bar and 2-D column charts.

You can set this property to a value from -100 through 100. If this property is set to -100, bars are positioned so that there is one bar width between them. If the overlap is 0 (zero), there is no space between bars (one bar starts immediately after the preceding bar). If the overlap is 100, bars are positioned on top of each other.


## Example

The following example sets the overlap for chart group one of the first chart in the active document to -50. You should run the example on a 2-D column chart that has two or more series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).Overlap = -50 
 End If 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

