---
title: Chart.ChartArea Property (Word)
keywords: vbawd10.chm79364157
f1_keywords:
- vbawd10.chm79364157
ms.prod: word
api_name:
- Word.Chart.ChartArea
ms.assetid: b16d78c0-7663-3ef9-c17a-02e7a024b344
ms.date: 06/08/2017
---


# Chart.ChartArea Property (Word)

Returns the complete chart area for the chart. Read-only  **[ChartArea](chartarea-object-word.md)** .


## Syntax

 _expression_ . **ChartArea**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example sets the chart area interior color of the first chart in the active document to red and sets the border color to blue.


```vb
With ActiveDocument.InlineShapes(1).Chart.ChartArea 
 .Interior.ColorIndex = 3 
 .Border.ColorIndex = 5 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

