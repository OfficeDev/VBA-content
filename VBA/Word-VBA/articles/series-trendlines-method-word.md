---
title: Series.Trendlines Method (Word)
keywords: vbawd10.chm123732122
f1_keywords:
- vbawd10.chm123732122
ms.prod: word
api_name:
- Word.Series.Trendlines
ms.assetid: 300dca01-097f-8a3d-4f63-a1841a92098e
ms.date: 06/08/2017
---


# Series.Trendlines Method (Word)

Returns a collection of all the trendlines for the series.


## Syntax

 _expression_ . **Trendlines**( **_Index_** )

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


### Return Value

A  **[Trendlines](trendlines-object-word.md)** object that represents all the treadlines for the series.


## Example

The following example adds a linear trendline to series one for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Trendlines.Add Type:=xlLinear 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

