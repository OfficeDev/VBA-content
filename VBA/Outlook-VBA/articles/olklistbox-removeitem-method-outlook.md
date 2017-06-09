---
title: OlkListBox.RemoveItem Method (Outlook)
keywords: vbaol11.chm1000277
f1_keywords:
- vbaol11.chm1000277
ms.prod: outlook
api_name:
- Outlook.OlkListBox.RemoveItem
ms.assetid: fe7bc0c4-d607-e4d1-b304-48b08f9c1e7a
ms.date: 06/08/2017
---


# OlkListBox.RemoveItem Method (Outlook)

Removes the specified item from the list.


## Syntax

 _expression_ . **RemoveItem**( **_Index_** )

 _expression_ A variable that represents an **OlkListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value indexing into the array of items in the list.|

## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](olklistbox-listcount-property-outlook.md)** -1), then an out-of-bounds error will be returned.


## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

