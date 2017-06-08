---
title: ChartArea.Clear Method (Word)
keywords: vbawd10.chm160039023
f1_keywords:
- vbawd10.chm160039023
ms.prod: word
api_name:
- Word.ChartArea.Clear
ms.assetid: 779f38e5-4a0e-fb1e-705b-d7f2c678a933
ms.date: 06/08/2017
---


# ChartArea.Clear Method (Word)

Clears the entire object.


## Syntax

 _expression_ . **Clear**

 _expression_ A variable that represents a **[ChartArea](chartarea-object-word.md)** object.


## Example

The following example clears the chart area (the chart data and formatting) of the first chart in the active document.


```vb
With ActiveDocument.InlineGroups(1) 
 If .HasChart Then 
 .Chart.ChartArea.Clear 
 End If 
End With
```


## See also


#### Concepts


[ChartArea Object](chartarea-object-word.md)

