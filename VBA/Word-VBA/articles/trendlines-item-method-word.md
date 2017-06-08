---
title: Trendlines.Item Method (Word)
keywords: vbawd10.chm102367232
f1_keywords:
- vbawd10.chm102367232
ms.prod: word
api_name:
- Word.Trendlines.Item
ms.assetid: 2aa9492d-efbb-155c-6836-cd1ac676e726
ms.date: 06/08/2017
---


# Trendlines.Item Method (Word)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **[Trendlines](trendlines-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The index number for the object.|

### Return Value

A  **[Trendline](trendline-object-word.md)** object that the collection contains.


## Example

The following example sets the number of units that the trendline on the first chart in the active document extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Trendlines.Item(1) 
 .Forward = 5 
 .Backward = .5 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Trendlines Object](trendlines-object-word.md)

