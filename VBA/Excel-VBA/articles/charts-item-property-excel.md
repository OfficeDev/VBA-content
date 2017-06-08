---
title: Charts.Item Property (Excel)
keywords: vbaxl10.chm217076
f1_keywords:
- vbaxl10.chm217076
ms.prod: excel
api_name:
- Excel.Charts.Item
ms.assetid: 792e3562-7d70-4356-7072-fa09cb40ec47
ms.date: 06/08/2017
---


# Charts.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Charts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example sets the number of units that the trendline on Chart1 extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
With Charts.Item("Chart1").SeriesCollection(1).Trendlines(1) 
 .Forward = 5 
 .Backward = .5 
End With
```


## See also


#### Concepts


[Charts Collection](charts-object-excel.md)

