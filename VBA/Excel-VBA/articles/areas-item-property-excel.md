---
title: Areas.Item Property (Excel)
keywords: vbaxl10.chm197074
f1_keywords:
- vbaxl10.chm197074
ms.prod: excel
api_name:
- Excel.Areas.Item
ms.assetid: 92b544c5-9122-f4d6-f22a-f5b21c7d6596
ms.date: 06/08/2017
---


# Areas.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **Areas** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|

## Example

This example clears the first area in the current selection if the selection contains more than one area.


```vb
If Selection.Areas.Count <> 1 Then 
 Selection.Areas.Item(1).Clear 
End If
```


## See also


#### Concepts


[Areas Collection](areas-object-excel.md)

