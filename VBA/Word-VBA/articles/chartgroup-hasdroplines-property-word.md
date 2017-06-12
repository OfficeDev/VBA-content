---
title: ChartGroup.HasDropLines Property (Word)
keywords: vbawd10.chm263454730
f1_keywords:
- vbawd10.chm263454730
ms.prod: word
api_name:
- Word.ChartGroup.HasDropLines
ms.assetid: 34743dd3-73f6-d125-a240-23984d31fa47
ms.date: 06/08/2017
---


# ChartGroup.HasDropLines Property (Word)

 **True** if the line chart or area chart has drop lines. Read/write **Boolean** .


## Syntax

 _expression_ . **HasDropLines**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

This property applies only to line and area charts. 


## Example

The following example enables drop lines for chart group one of the first chart in the active document and then sets their line style, weight, and color. You should run the example on a 2-D line chart that has one series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasDropLines = True 
 With .DropLines.Border 
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

