---
title: SeriesCollection.Item Method (Word)
keywords: vbawd10.chm150405120
f1_keywords:
- vbawd10.chm150405120
ms.prod: word
api_name:
- Word.SeriesCollection.Item
ms.assetid: 28793a84-8afe-ba65-7264-baf57e6b72ae
ms.date: 06/08/2017
---


# SeriesCollection.Item Method (Word)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **[SeriesCollection](seriescollection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **[Series](series-object-word.md)** object contained by the collection.


## Example

The following example sets the number of units that the trendline on the first chart in the active document extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection.Item(1).Trendlines.Item(1) 
 .Forward = 5 
 .Backward = .5 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[SeriesCollection Object](seriescollection-object-word.md)

