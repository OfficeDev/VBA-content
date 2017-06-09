---
title: EventList.Item Property (Visio)
keywords: vis_sdr.chm12713765
f1_keywords:
- vis_sdr.chm12713765
ms.prod: visio
api_name:
- Visio.EventList.Item
ms.assetid: 190cadcb-d985-30e9-eb9e-fdaa6a29aef4
ms.date: 06/08/2017
---


# EventList.Item Property (Visio)

Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **EventList** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Integer**|The index number of the object in its collection.|

### Return Value

Event


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index)
```


