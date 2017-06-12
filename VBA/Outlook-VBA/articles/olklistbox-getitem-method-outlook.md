---
title: OlkListBox.GetItem Method (Outlook)
keywords: vbaol11.chm1000268
f1_keywords:
- vbaol11.chm1000268
ms.prod: outlook
api_name:
- Outlook.OlkListBox.GetItem
ms.assetid: 23c47ede-8b72-e30a-b59a-1aa722be2064
ms.date: 06/08/2017
---


# OlkListBox.GetItem Method (Outlook)

Obtains a  **String** that represents an item at the specified location in the list.


## Syntax

 _expression_ . **GetItem**( **_Index_** )

 _expression_ A variable that represents an **OlkListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value that specifies the location of an item in the list.|

### Return Value

A  **String** value that represents the item at the specified location in the list.


## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](olklistbox-listcount-property-outlook.md)** -1), then an out-of-bounds error will be returned.


## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

