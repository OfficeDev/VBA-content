---
title: Windows.Item Property (Visio)
keywords: vis_sdr.chm11713765
f1_keywords:
- vis_sdr.chm11713765
ms.prod: visio
api_name:
- Visio.Windows.Item
ms.assetid: 61a17578-83c2-ce4e-95a4-739b32c7ad95
ms.date: 06/08/2017
---


# Windows.Item Property (Visio)

Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Windows** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Integer**|The index number of the object in its collection.|

### Return Value

Window


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index)
```


