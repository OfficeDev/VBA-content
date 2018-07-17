---
title: OlkComboBox.AddItem Method (Outlook)
keywords: vbaol11.chm1000230
f1_keywords:
- vbaol11.chm1000230
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.AddItem
ms.assetid: 8670b0ba-b715-e00d-0eb9-fa7279ae52b7
ms.date: 06/08/2017
---


# OlkComboBox.AddItem Method (Outlook)

Adds an item to the list, optionally specifying an index for the new item to appear in the list.


## Syntax

 _expression_ . **AddItem**( **_ItemText_** , **_Index_** )

 _expression_ A variable that represents an **OlkComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ItemText_|Required| **String**|Value to be added to the list in the combo box.|
| _Index_|Optional| **Long**|A 0-based value that specifies the order of the new item in the list.|

## Remarks

If the value of  _Index_ is equal to or larger than the number of elements in the list, the new item will be added to the end of the list.


## See also


#### Concepts


[OlkComboBox Object](olkcombobox-object-outlook.md)

