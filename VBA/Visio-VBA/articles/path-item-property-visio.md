---
title: Path.Item Property (Visio)
keywords: vis_sdr.chm15413765
f1_keywords:
- vis_sdr.chm15413765
ms.prod: visio
api_name:
- Visio.Path.Item
ms.assetid: e5e6e496-1573-375f-7e86-4b2c8b37a2a7
ms.date: 06/08/2017
---


# Path.Item Property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections, and for the **Path** and **Selection** objects. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Path** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|Contains the index of the object to retrieve.|

### Return Value

Curve


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index )
```


