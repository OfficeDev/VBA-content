---
title: Trendlines.Item Method (Excel)
keywords: vbaxl10.chm592076
f1_keywords:
- vbaxl10.chm592076
ms.prod: excel
api_name:
- Excel.Trendlines.Item
ms.assetid: e2bbc0fc-a618-8c84-f001-c77c0206cbf9
ms.date: 06/08/2017
---


# Trendlines.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Trendlines** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The index number for the object.|

### Return Value

A  **[Trendline](trendline-object-excel.md)** object contained by the collection.


## Example

This example sets the number of units that the trendline on Chart1 extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
With Charts("Chart1").SeriesCollection(1).Trendlines.Item(1) 
 .Forward = 5 
 .Backward = .5 
End With
```


## See also


#### Concepts


[Trendlines Object](trendlines-object-excel.md)

