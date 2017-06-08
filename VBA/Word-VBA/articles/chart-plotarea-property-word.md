---
title: Chart.PlotArea Property (Word)
keywords: vbawd10.chm79364154
f1_keywords:
- vbawd10.chm79364154
ms.prod: word
api_name:
- Word.Chart.PlotArea
ms.assetid: 440f7d57-c681-098e-45d6-a2f7aca6de07
ms.date: 06/08/2017
---


# Chart.PlotArea Property (Word)

Returns the plot area of a chart. Read-only  **[PlotArea](plotarea-object-word.md)** .


## Syntax

 _expression_ . **PlotArea**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example sets the color of the plot area interior for the first chart in the active document to cyan.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.PlotArea.Interior.ColorIndex = 8 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

