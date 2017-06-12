---
title: Chart.ChartData Property (Word)
keywords: vbawd10.chm79364189
f1_keywords:
- vbawd10.chm79364189
ms.prod: word
api_name:
- Word.Chart.ChartData
ms.assetid: d8234dd3-148f-b69a-8a4e-f22474080eab
ms.date: 06/08/2017
---


# Chart.ChartData Property (Word)

Returns information about the linked or embedded data associated with a chart. Read-only  **[ChartData](chartdata-object-word.md)** .


## Syntax

 _expression_ . **ChartData**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example uses the  **[Activate](chartdata-activate-method-word.md)** method to display the data associated with the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1).Chart.ChartData 
 .Activate 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

