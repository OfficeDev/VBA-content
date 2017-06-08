---
title: ChartGroup.DownBars Property (Word)
keywords: vbawd10.chm263454724
f1_keywords:
- vbawd10.chm263454724
ms.prod: word
api_name:
- Word.ChartGroup.DownBars
ms.assetid: ee556f66-cce6-aa8d-a837-ee8b0b93ba89
ms.date: 06/08/2017
---


# ChartGroup.DownBars Property (Word)

Returns the down bars on a line chart. Read-only  **[DownBars](downbars-object-word.md)** .


## Syntax

 _expression_ . **DownBars**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

This property applies only to line charts. 


## Example

The following example enables up bars and down bars for chart group one of the first chart in the active document and then sets their colors. You should run the example on a 2-D line chart that has two series that cross each other at one or more data points.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With Chart.ChartGroups(1) 
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

