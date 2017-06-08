---
title: GraphicItems.Item Property (Visio)
keywords: vis_sdr.chm16813765
f1_keywords:
- vis_sdr.chm16813765
ms.prod: visio
api_name:
- Visio.GraphicItems.Item
ms.assetid: bcd5ed67-3913-41ea-0d51-30ad24d04196
ms.date: 06/08/2017
---


# GraphicItems.Item Property (Visio)

Returns the  **GraphicItem** object at the specified index position in the **GraphicItems** collection. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **GraphicItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index of the object to retrieve.|

### Return Value

GraphicItem


## Remarks

 **Item** is the default property of the **GraphicItems** collection.

 The **GraphicItems** collection is indexed starting with 1.

When you retrieve objects from a collection, you can omit  **Item** from the expression because it is the default property of all collections. The following statement is equivalent to the syntax example given above:




```
objectReturned = expression(Index)
```


