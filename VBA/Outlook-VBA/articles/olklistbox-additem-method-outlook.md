---
title: OlkListBox.AddItem Method (Outlook)
keywords: vbaol11.chm1000276
f1_keywords:
- vbaol11.chm1000276
ms.prod: outlook
api_name:
- Outlook.OlkListBox.AddItem
ms.assetid: 0249eacc-746a-52bd-dcd3-fd25c96a5512
ms.date: 06/08/2017
---


# OlkListBox.AddItem Method (Outlook)

Adds an item to the list, optionally specifying an index for the new item to appear in the list.


## Syntax

 _expression_ . **AddItem**( **_ItemText_** , **_Index_** )

 _expression_ A variable that represents an **OlkListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ItemText_|Required| **String**|Value to be added to the list in the list box control.|
| _Index_|Optional| **Long**|A 0-based value that specifies the order of the new item in the list.|

## Remarks

If the value of  _Index_ is equal to or larger than the number of elements in the list, the new item will be added to the end of the list.


## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

