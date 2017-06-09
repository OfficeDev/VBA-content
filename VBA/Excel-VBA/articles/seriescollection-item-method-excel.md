---
title: SeriesCollection.Item Method (Excel)
keywords: vbaxl10.chm580077
f1_keywords:
- vbaxl10.chm580077
ms.prod: excel
api_name:
- Excel.SeriesCollection.Item
ms.assetid: 9a1f393b-e0b0-0887-b76e-471982ae0414
ms.date: 06/08/2017
---


# SeriesCollection.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **SeriesCollection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **[Series](series-object-excel.md)** object contained by the collection.


## Example

This example sets the number of units that the trendline on Chart1 extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
With Charts("Chart1").SeriesCollection.Item(1).Trendlines.Item(1) 
 .Forward = 5 
 .Backward = .5 
End With
```


## See also


#### Concepts


[SeriesCollection Object](seriescollection-object-excel.md)

