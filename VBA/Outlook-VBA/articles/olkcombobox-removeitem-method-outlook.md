---
title: OlkComboBox.RemoveItem Method (Outlook)
keywords: vbaol11.chm1000231
f1_keywords:
- vbaol11.chm1000231
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.RemoveItem
ms.assetid: 3fb8d3b4-3568-0b33-0672-8cb4cea31df2
ms.date: 06/08/2017
---


# OlkComboBox.RemoveItem Method (Outlook)

Removes the specified item from the list.


## Syntax

 _expression_ . **RemoveItem**( **_Index_** )

 _expression_ A variable that represents an **OlkComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value indexing into the array of items in the list.|

## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](olkcombobox-listcount-property-outlook.md)** -1), then an out-of-bounds error will be returned.


## See also


#### Concepts


[OlkComboBox Object](olkcombobox-object-outlook.md)

