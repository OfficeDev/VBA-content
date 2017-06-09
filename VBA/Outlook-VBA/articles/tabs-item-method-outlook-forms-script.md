---
title: Tabs.Item Method (Outlook Forms Script)
keywords: olfm10.chm2000310
f1_keywords:
- olfm10.chm2000310
ms.prod: outlook
ms.assetid: 3ceaf249-e2e8-4ef2-96f8-6379fbb81c4a
ms.date: 06/08/2017
---


# Tabs.Item Method (Outlook Forms Script)

Returns a member of a collection, either by position or by name.


## Syntax

 _expression_. **Item**( **_varg_**)

 _expression_A variable that represents a  **Tabs** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|varg|Required| **Variant**|A member's name or index within a collection.|

### Return Value

An Object that corresponds to the specified member in the collection.


## Remarks

The  _varg_ can be either a **String** or an **Integer**. If it is a  **String**, it must be a valid member name. If it is an  **Integer**, the minimum value is 0 and the maximum value is one less than the number of items in the collection.

If an invalid index or name is specified, an error occurs.


