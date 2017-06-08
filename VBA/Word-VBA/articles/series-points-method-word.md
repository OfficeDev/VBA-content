---
title: Series.Points Method (Word)
keywords: vbawd10.chm123732038
f1_keywords:
- vbawd10.chm123732038
ms.prod: word
api_name:
- Word.Series.Points
ms.assetid: 31f5763b-fdb9-de54-aff7-6fb3dc540a53
ms.date: 06/08/2017
---


# Series.Points Method (Word)

Returns a collection of all the points in the series.


## Syntax

 _expression_ . **Points**( **_Index_** )

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the point.|

### Return Value

A  **[Points](points-object-word.md)** object that represents all the points in the series.


## Example

The following example applies a data label to point one in series one of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Points(1).ApplyDataLabels 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

