---
title: ChartColorFormat.RGB Property (Word)
keywords: vbawd10.chm12059679
f1_keywords:
- vbawd10.chm12059679
ms.prod: word
api_name:
- Word.ChartColorFormat.RGB
ms.assetid: cd662ac4-e9ec-a6df-7af5-6d1fd13f86eb
ms.date: 06/08/2017
---


# ChartColorFormat.RGB Property (Word)

Returns the red-green-blue value of the specified color. Read-only  **Long** .


## Syntax

 _expression_ . **RGB**

 _expression_ A variable that represents a **[ChartColorFormat](chartcolorformat-object-word.md)** object.


## Example

The following example enables up and down bars, then adds a criss-cross pattern to the down bars and sets the pattern color to the chart area foreground fill color, for the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .ChartGroups(1).HasUpDownBars = True 
 .ChartGroups(1).DownBars.Interior.Pattern = xlPatternCrissCross 
 .ChartGroups(1).DownBars.Interior.PatternColor = _ 
 .ChartArea.Fill.ForeColor.RGB 
 End With 
 End If 
End With
```


## See also


#### Concepts


[ChartColorFormat Object](chartcolorformat-object-word.md)

