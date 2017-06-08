---
title: Interior.PatternColor Property (Word)
keywords: vbawd10.chm2818056
f1_keywords:
- vbawd10.chm2818056
ms.prod: word
api_name:
- Word.Interior.PatternColor
ms.assetid: 131f0006-6ed3-78f3-4888-8a3f47aeec78
ms.date: 06/08/2017
---


# Interior.PatternColor Property (Word)

Returns or sets the color of the interior pattern as an RGB value. Read/write  **Variant** .


## Syntax

 _expression_ . **PatternColor**

 _expression_ A variable that represents an **[Interior](interior-object-word.md)** object.


## Example

The following example enables up and down bars, then adds a criss-cross pattern to the down bars and sets the pattern color to blue, for the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.Pattern = xlPatternCrissCross 
 .DownBars.Interior.PatternColor = RGB(0, 0, 255) 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Interior Object](interior-object-word.md)

