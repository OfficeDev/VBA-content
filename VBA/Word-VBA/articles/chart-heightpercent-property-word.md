---
title: Chart.HeightPercent Property (Word)
keywords: vbawd10.chm79364106
f1_keywords:
- vbawd10.chm79364106
ms.prod: word
api_name:
- Word.Chart.HeightPercent
ms.assetid: b05873d9-a7b9-8980-28e7-057a90f7bb94
ms.date: 06/08/2017
---


# Chart.HeightPercent Property (Word)

Returns or sets the height of a 3-D chart as a percentage of the chart width (from 5 through 500 percent). Read/write  **Long** .


## Syntax

 _expression_ . **HeightPercent**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example sets the height of the first chart in the active document to 80 percent of its width. You should run the example on a 3-D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.HeightPercent = 80 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

