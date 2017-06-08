---
title: ChartGroup.UpBars Property (Word)
keywords: vbawd10.chm263454751
f1_keywords:
- vbawd10.chm263454751
ms.prod: word
api_name:
- Word.ChartGroup.UpBars
ms.assetid: 8581ad5f-94a1-0e12-3880-14ce2a7e9f03
ms.date: 06/08/2017
---


# ChartGroup.UpBars Property (Word)

Returns the up bars on a line chart. Read-only  **[UpBars](upbars-object-word.md)** .


## Syntax

 _expression_ . **UpBars**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

This property applies only to line charts.


## Example

The following example enables up and down bars for chart group one of the first chart in the active document, and then sets their colors. You should run the example on a 2-D line chart that contains two series that cross each other at one or more data points.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
 End With 
 End If 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

