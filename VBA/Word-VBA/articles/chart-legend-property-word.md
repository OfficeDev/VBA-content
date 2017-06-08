---
title: Chart.Legend Property (Word)
keywords: vbawd10.chm79364180
f1_keywords:
- vbawd10.chm79364180
ms.prod: word
api_name:
- Word.Chart.Legend
ms.assetid: b1ffdbfb-854c-bd65-dd63-d3b8d0547f67
ms.date: 06/08/2017
---


# Chart.Legend Property (Word)

Returns the legend for the chart. Read-only  **[Legend](legend-object-word.md)** .


## Syntax

 _expression_ . **Legend**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example enables the legend for the first chart in the active document and then sets the legend font color to blue.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .HasLegend = True 
 .Legend.Font.ColorIndex = 5 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

