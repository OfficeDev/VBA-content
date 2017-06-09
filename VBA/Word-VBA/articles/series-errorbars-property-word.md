---
title: Series.ErrorBars Property (Word)
keywords: vbawd10.chm123732127
f1_keywords:
- vbawd10.chm123732127
ms.prod: word
api_name:
- Word.Series.ErrorBars
ms.assetid: f3a4ecb9-2dd2-6d71-b5ca-8e1a3d47cd72
ms.date: 06/08/2017
---


# Series.ErrorBars Property (Word)

Returns the error bars for the series. Read-only  **[ErrorBars](errorbars-object-word.md)** .


## Syntax

 _expression_ . **ErrorBars**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Example

The following example sets the error bar color for series one of the first chart in the active document. You should run the example on a 2-D line chart that has error bars for series one.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1) 
 .ErrorBars.Border.ColorIndex = 8 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[Series Object](series-object-word.md)

