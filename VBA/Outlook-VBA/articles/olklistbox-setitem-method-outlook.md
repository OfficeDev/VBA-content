---
title: OlkListBox.SetItem Method (Outlook)
keywords: vbaol11.chm1000269
f1_keywords:
- vbaol11.chm1000269
ms.prod: outlook
api_name:
- Outlook.OlkListBox.SetItem
ms.assetid: 95232643-c547-f553-1d92-0f3fead18de9
ms.date: 06/08/2017
---


# OlkListBox.SetItem Method (Outlook)

Sets the item at the specified location in the list to the specified value.


## Syntax

 _expression_ . **SetItem**( **_Index_** , **_Item_** )

 _expression_ A variable that represents an **OlkListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value that specifies the location of an item in the list.|
| _Item_|Required| **String**|The value to be used to update the list at the specified location.|

## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](olklistbox-listcount-property-outlook.md)** -1), then an out-of-bounds error will be returned.


## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

