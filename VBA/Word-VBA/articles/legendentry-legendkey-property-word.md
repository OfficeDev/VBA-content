---
title: LegendEntry.LegendKey Property (Word)
keywords: vbawd10.chm4784302
f1_keywords:
- vbawd10.chm4784302
ms.prod: word
api_name:
- Word.LegendEntry.LegendKey
ms.assetid: 11aa8dfa-fdb9-d7f1-3c03-17ce68dcdbec
ms.date: 06/08/2017
---


# LegendEntry.LegendKey Property (Word)

Returns the legend key that is associated with the entry. Read-only  **[LegendKey](legendkey-object-word.md)** .


## Syntax

 _expression_ . **LegendKey**

 _expression_ A variable that represents a **[LegendEntry](legendentry-object-word.md)** object.


## Example

The following example sets the legend key for legend entry one on the first chart in the active document to be a triangle. You should run the example on a 2-D line chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Legend.LegendEntries(1).LegendKey _ 
 .MarkerStyle = xlMarkerStyleTriangle 
 End If 
End With
```


## See also


#### Concepts


[LegendEntry Object](legendentry-object-word.md)

