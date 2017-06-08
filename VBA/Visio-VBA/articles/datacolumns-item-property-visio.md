---
title: DataColumns.Item Property (Visio)
keywords: vis_sdr.chm16613765
f1_keywords:
- vis_sdr.chm16613765
ms.prod: visio
api_name:
- Visio.DataColumns.Item
ms.assetid: c61db4d2-a802-9e02-991e-af0fb9783989
ms.date: 06/08/2017
---


# DataColumns.Item Property (Visio)

Returns the  **DataColumn** object at the specified index position, or of the specified name, in the **DataColumns** collection. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Item**( **_IndexOrName_** )

 _expression_ A variable that represents a **DataColumns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _IndexOrName_|Required| **Variant**|The index (as a  **Long** ) or the name (as a **String** ) of the data column to retrieve.|

### Return Value

DataColumn


## Remarks

 **Item** is the default property of the **DataColumns** collection.

When you retrieve objects from a collection, you can omit  **Item** from the expression because it is the default property of all collections. The following statements are equivalent to the syntax example given above:




```
objectReturned = expression(Index) 
objectReturned = expression(Name) 

```

 The **DataColumns** collection is indexed starting with 1.


