---
title: AccelItems.Item Property (Visio)
keywords: vis_sdr.chm14613765
f1_keywords:
- vis_sdr.chm14613765
ms.prod: visio
api_name:
- Visio.AccelItems.Item
ms.assetid: c6ac3d03-4b13-141f-d1fd-dfbf671435fd
ms.date: 06/08/2017
---


# AccelItems.Item Property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_lIndex_** )

 _expression_ A variable that represents a **AccelItems** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|Contains the index of the object to retrieve.|

### Return Value

AccelItem


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index )
```


