---
title: Selection.Item Property (Visio)
keywords: vis_sdr.chm11113765
f1_keywords:
- vis_sdr.chm11113765
ms.prod: visio
api_name:
- Visio.Selection.Item
ms.assetid: 3f09566d-eec6-0c20-87bc-60db45d3e23f
ms.date: 06/08/2017
---


# Selection.Item Property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections, and for the **Path** and **Selection** objects. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|Contains the index of the object to retrieve.|

### Return Value

Shape


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index )
```


